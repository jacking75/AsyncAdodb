# AsyncAdodb
1 header file로 만든 C++ ADO 라이브러리이다.  
이 라이브러리는 [네이버의 온라인 서버 제작자 모임](https://cafe.naver.com/ongameserver/3412) 의 멤버인 김영찬님이 공개한 라이브러리를 수정한 버전이다.  

## 사용법  
설정하기  
```
SAdoConfig adoconfig;
adoconfig.SetIP(_T("localhost\\TEST"));   // 디비 주소
adoconfig.SetUserID(_T("sa"));	   // 아이디
adoconfig.SetPassword(_T("dev"));	   // 패스워드
adoconfig.SetInitCatalog(_T("TEST"));       // 디비 이름
adoconfig.SetConnectionTimeout(3);
adoconfig.SetRetryConnection(true);
adoconfig.SetMaxConnectionPool(3);

CAdoManager* pAdomanager = new CAdoManager( adoconfig );
```
    
일반적인 SQL 문 예  
```
// Users라는 테이블을 만들었고 ID와 PWD는 nvarchar(16) 입니다

CAdoManager* pAdomanager = new CAdoManager(adoconfig);
CAdo* pAdo = NULL;
{
    CScopedAdo scopedado(pAdo, pAdomanager, false);
    pAdo->SetQuery(_T("SELECT UID, PWD FROM Users WHERE ID='jacking'"));
    pAdo->Execute(adCmdText);
    if(!pAdo->IsSuccess())
    {
        std::wcout << L"쿼리문 실패" << std::endl;
        return 0;
    }

    int nUID = 0;
    WCHAR szPWD[16];
    if(!pAdo->GetEndOfFile() )
    {
       pAdo->GetFieldValue(_T("UID"), nUID);
       pAdo->GetFieldValue(_T("PWD"), szPWD, 16);
    }
    else
    {
       std::wcout << L"jacking은 없습니다" << std::endl;
       return 0;
    }

```
  
일반 INSERT 쿼리문 1  
```
CAdoManager* pAdomanager = new CAdoManager(adoconfig);
CAdo* pAdo = NULL;
{
   CScopedAdo scopedado(pAdo, pAdomanager, false);

   pAdo->SetQuery( _T("Insert Into Users Values( 'jacking2', '1111' )") );
   pAdo->Execute(adCmdText);
  
   if( !pAdo->IsSuccess() ) 
   {
       std::wcout << L"쿼리문 실패" << std::endl;
       return 0;
   }
   else
  {
       std::wcout << L"쿼리문 성공" << std::endl;
  }

```
  
일반 INSERT 쿼리문 1    
```
CAdoManager* pAdomanager = new CAdoManager(adoconfig);
CAdo* pAdo = NULL;
{
   CScopedAdo scopedado(pAdo, pAdomanager, true);  // 트랜잭션 설정

   pAdo->SetQuery( _T("Insert Into Users Values( 'jacking3', '1111' )") );
   pAdo->Execute(adCmdText);
  
   if( !pAdo->IsSuccess() ) 
   {
       std::wcout << L"쿼리문 실패" << std::endl;
       return 0;
   }
   else
  {
       std::wcout << L"쿼리문 성공" << std::endl;
  }

  pAdo->SetCommit(true);  // true에 의해서 commit 된다

```
  
  
## ADO 연결 문자열
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
- [ado ms-sql, mysql](http://sakbals.tistory.com/entry/%EC%B4%88%EA%B8%89-%EA%B0%84%EB%8B%A8%ED%95%98%EA%B2%8C-ADO-oledb-%EC%82%AC%EC%9A%A9%ED%95%98%EA%B8%B0)
- [C++로 Mysql을 ADO로 연결하려고 하는데 출력시 한글이 깨집니다](https://kldp.org/node/142606)
- [msado15.dll을 이용해서 MFC에서 ADO로 데이테베이스 억세스하기](http://wwwi.tistory.com/80)
- [(일어) Visual Studio で MySQL データベースに接続する](http://d.hatena.ne.jp/hilapon/20151215/1450166056)
- [How To Connect C++ To Mysql](https://www.dreamincode.net/forums/topic/88724-how-to-connect-c-to-mysql/)
- [(일어) Connector/ODBC アプリケーション別情報](http://download.nust.na/pub6/mysql/doc/refman/5.1/ja/myodbc-usagenotes-apptips.html)
- [(일어) 接続文字列一覧](http://kojama.la.coocan.jp/works/rdbms/conn/connstr.html)
- [Tutorial: Moving from MySQL to ADODB](http://web.unife.it/lib/adodb/docs/tute.htm)
- [ADO Connection Strings](https://www.codeproject.com/Articles/2304/ADO-Connection-Strings)


##
/*
pAdo->UpdateParameter(_T("@v_bbb"), (TCHAR*)tstrString.c_str(), tstrString.size()); 
동일 쿼리문을 여러번 실행 시킬 필요가 있을때 사용합니다. 예로 들어 여러 건수의 값을 insert할때
*/



/*
2008. 12. 10 카페에 올려주신 예제 코드
http://cafe.naver.com/ongameserver/3415

create table Users(
id int,
PassWord varchar(16),
Level smallint,
Money int
)


insert into Users values(5, "sdofjoaf", 255, 10100) 

 

//프로시저 입니다.

create procedure dbo.sp_user_authe
    @id int,
    @CharCD bigint output
as
begin
select [PassWord], level, money from Users where id = @id
set @CharCD = 2736485867678
return 77 
end

 

 

//소스

#include "stdafx.h"
#include "AdoManager.h"

int AdoWork();

CAdoManager* adomanager=NULL;

int _tmain(int argc, _TCHAR* argv[])
{  
#ifdef _UNICODE
  _tsetlocale(LC_ALL,_T("Korean"));
#endif

  SAdoConfig adoconfig;
  adoconfig.SetIP(_T("172.20.0.63\\MYDB"));    //디비주소
  adoconfig.SetUserID(_T("xxxxxxx"));          //아이디
  adoconfig.SetPassword(_T("xxxxxx)"));        //패스워드
  adoconfig.SetInitCatalog(_T("GAMEDB"));      //디비명
  adoconfig.SetConnectionTimeout(3);
  adoconfig.SetRetryConnection(true);
  adoconfig.SetMaxConnectionPool(20);

  adomanager = new CAdoManager(adoconfig);

  AdoWork();
  return 0;
}


int AdoWork()
{
  {
    CAdo* pAdo = NULL;
    //adomanager로부터 ado연결을 가져온다. 명시적 트랜잭션을 사용하지 않는다.
    CScopedAdo scopedado(pAdo, adomanager, false);

    INT nReturn = 0;
    INT nId = 5;
    INT64 i64CharCd = 0;
    pAdo->CreateParameter(_T("return"),adInteger, adParamReturnValue, nReturn);  //리턴값 얻기
    pAdo->CreateParameter(_T("id"),adInteger, adParamInput, nId);
    pAdo->CreateParameter(_T("CharCd"),adBigInt, adParamInputOutput, i64CharCd); //Int64 및 Output 파라메터로 설정

    pAdo->SetQuery(_T("GameDB..sp_user_authe"));
    pAdo->Execute();
    if(!pAdo->IsSuccess()){ return 0;  }

    TCHAR tszPassWorld[17] = {0, };
    WORD wLevel = 0;
    INT nMoney = 0;

    while(!pAdo->GetEndOfFile())
    {
      ISFAIL(pAdo->GetFieldValue(_T("PassWord"), tszPassWorld, 16));
      ISFAIL(pAdo->GetFieldValue(_T("Level"), wLevel));
      ISFAIL(pAdo->GetFieldValue(_T("Money"), nMoney));
      _tprintf(_T("%s %d, %d\n"), tszPassWorld, wLevel, nMoney);
      pAdo->MoveNext();
    }

    pAdo->GetParameter(_T("return"), nReturn);
    pAdo->GetParameter(_T("CharCd"), i64CharCd);
    _tprintf(_T("return:%d, CharCd:I64d\n"), nReturn, i64CharCd);
  }
}

*/

/*
무모님 소스에 있던 예제 코드

unsigned __stdcall AdoFunc(void *pArg);
int AdoWork();
int AdoWork2();

CAdoManager* adomanager=NULL;



int _tmain(int argc, _TCHAR* argv[])
{	
#ifdef _UNICODE
	_tsetlocale(LC_ALL,_T("Korean"));
#endif

	SAdoConfig adoconfig;
	adoconfig.SetIP(_T("127.0.0.1"));
//	adoconfig.SetDSN(_T("ADODB_CONN"));
	adoconfig.SetUserID(_T("adotester"));
	adoconfig.SetPassword(_T("adotester"));
	adoconfig.SetInitCatalog(_T("ADODB"));
	adoconfig.SetConnectionTimeout(3);
	adoconfig.SetRetryConnection(true);
	adoconfig.SetMaxConnectionPool(20);

	adomanager = new CAdoManager(adoconfig);

	HANDLE hThread;
	unsigned threadID;
	hThread = (HANDLE) _beginthreadex(NULL, 0, AdoFunc, NULL, 0, &threadID);

	while(1)
	{
		AdoWork();
	}

	WaitForSingleObject(hThread, INFINITE);
	CloseHandle(hThread);
	return 0;
}

unsigned __stdcall AdoFunc(void *pArg)
{
	while(1)
	{
		AdoWork2();
	}
	return 0;
}


int AdoWork()
{
	DWORD nParam = 1231415151, nRtnParam = 0;
	TCHAR tszParam[30] = _T("uuuuu");
	BYTE pbyParam[10000] = {0XFF, 0XFE, 0X00, 0X01, 0X06, 0X07,};
	BYTE pbyDATA[10000]={0XFF, 0XFE, 0X00, 0X01, 0X06, 0X07,};
	bool bBoolValue = 1;
	char byByteValue = 100;
	WORD wWordValue = 20000;
	int nReturn = 0;
	INT64 i64BigIntValue = 8223372036854775801;
	COleDateTime oleTime;
	float fValue = 74.234738123f;
	int nSize = 0;
	TCHAR tszdsjo[4]={0,};
	_variant_t vValue(1000);
	BYTE pValue[1000]={0,};
	oleTime.SetDateTime(2008, 2, 28, 23, 59, 59);


	{
		CAdo* pAdo = NULL;
		//adomanager로부터 ado연결을 가져온다. 명시적 트랜잭션을 사용하지 않는다.
		CScopedAdo scopedado(pAdo, adomanager, false);

		pAdo->CreateParameter(_T("return"),adInteger, adParamReturnValue, nReturn);
//		pAdo->CreateNullParameter(_T("@v_aaa"), adInteger, adParamInputOutput);
		pAdo->CreateParameter(_T("@v_aaa"),adInteger, adParamInputOutput, vValue);
		pAdo->CreateParameter(_T("@v_bbb"),adVarChar, adParamInputOutput, (TCHAR*)NULL, 1);
		pAdo->CreateParameter(_T("@v_ccc"),adVarBinary, adParamInputOutput, (BYTE*)NULL, 1);
		pAdo->CreateParameter(_T("@v_ddd"),adBoolean, adParamInputOutput, bBoolValue);
		pAdo->CreateParameter(_T("@v_eee"),adTinyInt, adParamInputOutput, byByteValue);
		pAdo->CreateParameter(_T("@v_fff"),adSmallInt, adParamInputOutput, wWordValue);
		pAdo->CreateParameter(_T("@v_ggg"),adBigInt, adParamInputOutput, i64BigIntValue);
		//pAdo->CreateParameter<COleDateTime&>(_T("@v_hhh"),adDBTimeStamp, adParamInputOutput, oleTime);
		//pAdo->CreateParameter<COleDateTime&>(_T("@v_iii"),adDBTimeStamp, adParamInputOutput, oleTime);
		pAdo->CreateParameter(_T("@v_hhh"),adDBTimeStamp, adParamInputOutput, oleTime);
		pAdo->CreateParameter(_T("@v_iii"),adDBTimeStamp, adParamInputOutput, oleTime);
		pAdo->CreateParameter(_T("@v_jjj"),adDouble, adParamInputOutput, fValue);
		pAdo->CreateParameter(_T("@v_kkk"),adBinary, adParamInputOutput, (BYTE*)NULL, 1);

		tstring tstrString;
		int i = 5;

		while(i > 0)
		{
//			pAdo->UpdateParameter(_T("@v_aaa"), i * 1000);
//			pAdo->UpdateNullParameter(_T("@v_aaa"));
			tstrString += _T("yy");
			oleTime.SetDateTime(2007, 12, 12+i, 23, 10+i, 1+i);
			pAdo->UpdateParameter<COleDateTime&>(_T("@v_hhh"), oleTime);
			pAdo->UpdateParameter(_T("@v_hhh"), oleTime);
			pAdo->UpdateParameter(_T("@v_bbb"), (TCHAR*)tstrString.c_str(), tstrString.size());
//			pAdo->UpdateParameter(_T("@v_bbb"), (TCHAR*)NULL, 1);
			pAdo->UpdateParameter(_T("@v_kkk"), pbyParam, 10);
			pAdo->UpdateParameter(_T("@v_ccc"), pbyParam, 10);
			//pAdo->UpdateParameter(_T("@v_ccc"), (BYTE*)NULL, i);
			if(!pAdo->IsSuccess()){ return 0;	}

			pAdo->SetQuery(_T("adotestproc"));

			pAdo->Execute();
			if(!pAdo->IsSuccess()){ return 0;	}

			nParam = 0;
			while(!pAdo->GetEndOfFile())
			{
				ISFAIL(pAdo->GetFieldValue(_T("aaa"), nParam));
				_tprintf(_T("%d\n"), nParam);
				pAdo->MoveNext();
			}
			if(!pAdo->IsSuccess()){ return 0;}

			pAdo->NextRecordSet();
			if(!pAdo->IsSuccess()){ return 0;}

			while(!pAdo->GetEndOfFile())
			{
				ISFAIL(pAdo->GetFieldValue(_T("bbb"), tszParam, 30));
				_tprintf(_T("bbb%d\n"), nParam);
				pAdo->MoveNext();
			}
			if(!pAdo->IsSuccess()){return 0;}


			pAdo->NextRecordSet();
			if(!pAdo->IsSuccess()){ return 0;}

			while(!pAdo->GetEndOfFile())
			{
				memset(pValue, 0x00, 50);
				ISFAIL(pAdo->GetFieldValue(_T("ccc"), pValue, 10, nSize));
				_tprintf(_T("Size(%u)"), nSize);
				for(int i = 0; i < nSize; ++i)
					_tprintf(_T("%02X"), pValue[i]);
				_tprintf(_T("\n"));
				pAdo->MoveNext();
			}
			if(!pAdo->IsSuccess()){return 0;}

			pAdo->NextRecordSet();
			if(!pAdo->IsSuccess()){ return 0;}

			while(!pAdo->GetEndOfFile())
			{
				ISFAIL(pAdo->GetFieldValue(_T("ddd"), bBoolValue));
				_tprintf(_T("%d\n"), bBoolValue);
				pAdo->MoveNext();
			}
			if(!pAdo->IsSuccess()){ return 0;}

			pAdo->NextRecordSet();
			if(!pAdo->IsSuccess()){ return 0;}

			while(!pAdo->GetEndOfFile())
			{
				ISFAIL(pAdo->GetFieldValue(_T("eee"), byByteValue));
				_tprintf(_T("%d\n"), byByteValue);
				pAdo->MoveNext();
			}
			if(!pAdo->IsSuccess()){ return 0;}

			pAdo->NextRecordSet();
			if(!pAdo->IsSuccess()){ return 0;}

			while(!pAdo->GetEndOfFile())
			{
				ISFAIL(pAdo->GetFieldValue(_T("fff"), wWordValue));
				_tprintf(_T("%d\n"), wWordValue);
				pAdo->MoveNext();
			}
			if(!pAdo->IsSuccess()){ return 0;}


			pAdo->NextRecordSet();
			if(!pAdo->IsSuccess()){ return 0;}

			while(!pAdo->GetEndOfFile())
			{
				ISFAIL(pAdo->GetFieldValue(_T("ggg"), i64BigIntValue));
				_tprintf(_T("%I64d\n"), i64BigIntValue);
				pAdo->MoveNext();
			}
			if(!pAdo->IsSuccess()){ return 0;}

			pAdo->NextRecordSet();
			if(!pAdo->IsSuccess()){ return 0;}

			while(!pAdo->GetEndOfFile())
			{
				ISFAIL(pAdo->GetFieldValue(_T("hhh"), oleTime));
				_tprintf(_T("%s\n"), oleTime.Format(_T("%Y-%m-%d %H:%M:%S")));
				pAdo->MoveNext();
			}
			if(!pAdo->IsSuccess()){ return 0;}

			pAdo->NextRecordSet();
			if(!pAdo->IsSuccess()){ return 0;}

			while(!pAdo->GetEndOfFile())
			{
				ISFAIL(pAdo->GetFieldValue(_T("iii"), oleTime));
				_tprintf(_T("%s\n"), oleTime.Format(_T("%Y-%m-%d %H:%M:%S")));
				pAdo->MoveNext();
			}
			if(!pAdo->IsSuccess()){ return 0;}

			pAdo->NextRecordSet();
			if(!pAdo->IsSuccess()){ return 0;}

			while(!pAdo->GetEndOfFile())
			{
				ISFAIL(pAdo->GetFieldValue(_T("jjj"), fValue));
				_tprintf(_T("%.10lf\n"), fValue);
				pAdo->MoveNext();
			}
			if(!pAdo->IsSuccess()){ return 0;}

			pAdo->GetParameter(_T("return"), nRtnParam);
			pAdo->GetParameter(_T("@v_aaa"), nParam);
			pAdo->GetParameter(_T("@v_bbb"), tszParam, 100);
			pAdo->GetParameter(_T("@v_ccc"), pbyDATA, 10, nSize);
			pAdo->GetParameter(_T("@v_ddd"), bBoolValue);
			pAdo->GetParameter(_T("@v_eee"), byByteValue);
			pAdo->GetParameter(_T("@v_fff"), wWordValue);
			pAdo->GetParameter(_T("@v_ggg"), i64BigIntValue);
			pAdo->GetParameter(_T("@v_hhh"), oleTime);
			pAdo->GetParameter(_T("@v_iii"), oleTime);
			pAdo->GetParameter(_T("@v_jjj"), fValue);
			pAdo->GetParameter(_T("@v_kkk"), pbyDATA, 10, nSize);
			if(!pAdo->IsSuccess()){return 0;}

			_tprintf(_T("\nReturn value: %d\n"), nRtnParam);
			_tprintf(_T("param out1: %d\n"), nParam); 
			_tprintf(_T("param out2: %s\n"), tszParam);
			_tprintf(_T("param out4: %d\n"), bBoolValue); 
			_tprintf(_T("param out5: %d\n"), byByteValue); 
			_tprintf(_T("param out6: %d\n"), wWordValue); 
			_tprintf(_T("param out7: %I64d\n"), i64BigIntValue);
			_tprintf(_T("param out8: %s\n"), oleTime.Format(_T("%Y-%m-%d %H:%M:%S")));
			_tprintf(_T("param out9: %.10lf\n"), fValue);
			_tprintf(_T("param out10: "), fValue);
			for(int k = 0; k < nSize; ++k)
			{
				_tprintf(_T("%02X"), pbyDATA[i]);
			}


			_tprintf(_T("\n"));

			i--;
		}
	}
	return 0;
}

int AdoWork2()
{
	CAdo* pAdo = NULL;
	{
		CScopedAdo scopedado(pAdo, adomanager, false);

		pAdo->SetQuery(_T("Insert Into AdoTest Values(999, 'sdjfojaf', 12314, 1, 25, 25, 25, '20071231', '20070101', 25.3, null)"));
		pAdo->Execute(adCmdText);
		if(!pAdo->IsSuccess()){return 0;}
		pAdo->SetQuery(_T("Insert Into AdoTest Values(888, 'sdjfojaf', 12314, 1, 25, 25, 25, '20071231', '20070101', 25.3, null)"));
		//		pAdo->SetQuery(_T("select aaa from adotest"));
		pAdo->Execute(adCmdText);
		if(!pAdo->IsSuccess()){return 0;}

		pAdo->SetCommit(false);


		//DWORD nParam = 0;
		//while(!pAdo->GetEndOfFile())
		//{
		//	ISFAIL(pAdo->GetFieldValue<DWORD>(_T("aaa"), nParam));
		//	_tprintf(_T("%d\n"), nParam);
		//	pAdo->MoveNext();
		//}
		//if(!pAdo->IsSuccess()){ return 0;}
	}
	return 0;
}
*/