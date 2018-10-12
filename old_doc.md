## 사용법  
### 설정하기  

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
  
  
### 일반적인 SQL 문 예  

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
  
  
### 일반 INSERT 쿼리문 1  

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
  
  
### 일반 INSERT 쿼리문 2    

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
   

### SP 예제  

```
create table Users(
id int,
PassWord varchar(16),
Level smallint,
Money int
)
```  
  
```  
insert into Users values(5, "sdofjoaf", 255, 10100) 
```
     
```
// 프로시저.
create procedure dbo.sp_user_authe
    @id int,
    @CharCD bigint output
as
begin
select [PassWord], level, money from Users where id = @id
set @CharCD = 2736485867678
return 77 
end
```
  
```
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
```  
