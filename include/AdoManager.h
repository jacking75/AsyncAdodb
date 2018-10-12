/*
김영찬님이 공개한 ADO 라이브러리를 수정한 것이다.
*/
#pragma once

// MS가 제공하는 ado 라이브러리가 있는 위치
#import "C:\Program Files\Common Files\System\ADO\msado15.dll" rename("EOF", "EndOfFile") no_namespace

#include <string>
#include <mutex>

#include <ole2.h>
#include <atlcomtime.h>


#define ISFAIL(a) if(!(a)) break

/** 
\brief		커넥션풀 ADO객체를 생성하여 stack에 관리한다.
\author		김영찬
*/
// 주석 참고 문헌
// http://msdn.microsoft.com/ko-kr/library/5ast78ax(v=vs.80).aspx
// http://msdn.microsoft.com/ko-kr/library/z04awywx(v=vs.80)

namespace AsyncAdodb
{
	/// <summary>
	/// DB 설정 정보 </summary>
	class DBConfig
	{
	public:
		DBConfig()
			:m_nConnectionTimeout(0),
			m_nCommandTimeout(0),
			m_bRetryConnection(false),
			m_bCanWriteErrorLog(false)
		{
		}

		/// <summary>
		/// 접속 정보 설정 </summary>
		void Setting(const std::wstring strDBAddress,
			const std::wstring strID,
			const std::wstring strPassWord,
			const std::wstring strDBName,
			const int nConnectionTimeOut,
			const bool bIsRetryConnection,
			const int nMaxConnectionPool,
			const bool bCanWriteErrorLog = true
		)
		{
			m_strDataSource = L";Data Source=";
			m_strDataSource += strDBAddress;
			SetProvider();

			m_strUserID = strID;
			m_strPassword = strPassWord;
			m_strInitialCatalog = strDBName;

			m_nConnectionTimeout = nConnectionTimeOut;
			m_bRetryConnection = bIsRetryConnection;
			m_nMaxConnectionPool = nMaxConnectionPool;

			m_bCanWriteErrorLog = bCanWriteErrorLog;
		}

		/// <summary>
		/// DB DSN 주소 설정 </summary>
		void SetDSN(wchar_t* pszString)
		{
			m_strDataSource.clear();
			m_strDSN = L";DSN=";
			m_strDSN += pszString;
		}
				
		/// <summary> 
		/// 명령 대기 시간 설정. SQL 명령을 내린 후 지정한 시간까지만 처리를 기다림 </summary>
		void SetCommandTimeout(int nCommendTimeout) { m_nCommandTimeout = nCommendTimeout; }

		std::wstring GetConnectionString()
		{
			if (m_strDataSource.empty())
			{
				m_strConnectingString = m_strDSN;
			}
			else
			{
				m_strConnectingString = m_strDataSource;
			}

			return m_strConnectingString;
		}

		std::wstring& GetUserID() { return m_strUserID; }

		std::wstring& GetPassword() { return m_strPassword; }

		std::wstring& GetInitCatalog() { return m_strInitialCatalog; }

		std::wstring& GetProvider() { return m_strProvider; }

		int GetConnectionTimeout() { return m_nConnectionTimeout; }

		bool IsCanRetryConnection() { return m_bRetryConnection; }

		int GetMaxConnectionPool() { return m_nMaxConnectionPool; }

		bool CanWriteErrorLog() { return m_bCanWriteErrorLog; }


	private:
		/// <summary> DB API 프로바이더 설정 </summary>
		void SetProvider(wchar_t* pString = L"SQLOLEDB.1")
		{
			m_strProvider = pString;
		}


		/// <summary> 연결 문자열 </summary>
		std::wstring m_strConnectingString;

		/// <summary> 사용할 데이터베이스 이름 </summary>
		std::wstring m_strInitialCatalog;

		/// <summary> 데이터베이스 주속 </summary>
		std::wstring m_strDataSource;

		/// <summary> DB 접속 아이디 </summary>
		std::wstring m_strUserID;

		/// <summary> DB 접속 패스워드 </summary>
		std::wstring m_strPassword;

		/// <summary> DB API 프로바이더 </summary>
		std::wstring m_strProvider;

		/// <summary> DB의 DSN 주소 </summary>
		std::wstring m_strDSN;

		/// <summary>  </summary>
		int m_nConnectionTimeout;

		/// <summary> ?? </summary>
		int m_nCommandTimeout;

		/// <summary> 재연결 여부 </summary>
		bool m_bRetryConnection;

		/// <summary> 연결 풀의 최대 개수 </summary>
		int m_nMaxConnectionPool;

		/// <summary> 에러 로그 출력 가능 여부 </summary>
		bool m_bCanWriteErrorLog;
	};


	//<< 동기화 객체들 정의 >>
	const INT32 MAX_SPIN_LOCK_COUNT = 4000;

	class ISynchronizeObj
	{
	public:
		virtual void Lock() = 0;
		virtual void UnLock() = 0;
	};


	/// <summary>
	/// Win32 API의 스핀락 크리티컬섹션 </summary>
	class CSSpinLockWin32 : public ISynchronizeObj
	{
	public:
		CSSpinLockWin32()
		{
			InitializeCriticalSectionAndSpinCount(&m_lock, MAX_SPIN_LOCK_COUNT);
		}

		~CSSpinLockWin32()
		{
			DeleteCriticalSection(&m_lock);
		}

		virtual void Lock() { EnterCriticalSection(&m_lock); }
		virtual void UnLock() { LeaveCriticalSection(&m_lock); }

	private:
		CRITICAL_SECTION m_lock;
	};


	/// <summary>
	/// C++11의 뮤텍스 사용(Windows 환경에서는 Win32 API 크리티컬섹션 사용) </summary>
	class StandardLock : public ISynchronizeObj
	{
	public:
		StandardLock() {}
		~StandardLock() {}

		virtual void Lock() { m_lock.lock(); }
		virtual void UnLock() { m_lock.unlock(); }

	private:
		std::mutex m_lock;
	};


	/// <summary>
	/// 락을 객체 생성과 해제에서 자동으로 락과 언락을 하도록 동작 </summary>
	class ScopedLock
	{
	public:
		ScopedLock(ISynchronizeObj &SyncObj) :m_SyncObj(SyncObj)
		{
			m_SyncObj.Lock();
		}

		~ScopedLock()
		{
			m_SyncObj.UnLock();
		}

	private:
		ISynchronizeObj &m_SyncObj;
	};


	class AdoDB
	{
	public:
		AdoDB(DBConfig& adoconfig) :m_bAutoCommit(false),
			m_Config(adoconfig),
			m_pConnection(nullptr),
			m_pCommand(nullptr),
			m_pRecordset(nullptr),

			m_bCanGetParamGetFiled(true),
			m_bCanCommitTransaction(true),
			m_strParameterName(),
			m_strColumnName(),
			m_strQuery(),
			m_strCommand()
		{
			if (FAILED(::CoInitialize(nullptr)))
			{
				LOG(L"::CoInitialize Fail!!");
				return;
			}

			m_pConnection.CreateInstance(__uuidof(Connection));
			m_pCommand.CreateInstance(__uuidof(Command));
		}

		~AdoDB()
		{
			Close();
		}

		/// <summary>
		/// 초기화 - 연결풀에서 재사용하기 위해 이곳에서 초기화 시켜줌 </summary>
		void Init()
		{
			m_bAutoCommit = false;
			m_bCanGetParamGetFiled = true;
			m_strParameterName.clear();
			m_strColumnName.clear();
			m_strQuery.clear();
			m_strCommand.clear();
		}

		/// <summary>
		/// 연결 설정 - IP 및 DSN 접속 
		/// <param name="CursorLocation"> 배치 작업일 경우 adUseClientBatch 옵션 사용 </param>
		/// <returns> 성공(TRUE) 실패(FLASE) </returns>
		/// </summary>
		bool Open(CursorLocationEnum CursorLocation = adUseClient)
		{
			m_strCommand = L"Open()";

			try
			{
				if (m_Config.GetConnectionTimeout() != 0) {
					m_pConnection->PutConnectionTimeout(m_Config.GetConnectionTimeout());
				}

				m_pConnection->CursorLocation = CursorLocation;

				if (!m_Config.GetProvider().empty()) { //ip접속일 경우 Provider 사용
					m_pConnection->put_Provider((_bstr_t)m_Config.GetProvider().c_str());
				}

				m_pConnection->Open((_bstr_t)m_Config.GetConnectionString().c_str(), 
							(_bstr_t)m_Config.GetUserID().c_str(),
							(_bstr_t)m_Config.GetPassword().c_str(), NULL);

				if (m_pConnection->GetState() == adStateOpen) {
					m_pConnection->DefaultDatabase = m_Config.GetInitCatalog().c_str();
				}

				m_pCommand->ActiveConnection = m_pConnection;
				return true;
			}
			catch (_com_error &e)
			{
				dump_com_error(e);
				return false;
			}
		}

		/// <summary>
		/// 재연결 옵션이 있을 경우 재연결 시도 </summary>
		bool RetryOpen()
		{
			if (m_pConnection->GetState() != adStateClosed) {
				return false;
			}

			return Open();
		}

		/// <summary>
		/// 연결 종료 </summary>
		void Close()
		{
			if (m_pConnection == nullptr) {
				return;
			}

			try
			{
				if (m_pConnection->GetState() != adStateClosed) {
					m_pConnection->Close();
				}
			}
			catch (...)
			{

			}
		}

		/// <summary>
		///  커넥션풀에서 재사용하기 위한 커맨드 객체 재생성 </summary>
		void Release()
		{
			m_pCommand.Release();
			m_pCommand.CreateInstance(__uuidof(Command));

			if (m_pConnection->GetState() != adStateClosed) {
				m_pCommand->ActiveConnection = m_pConnection;
			}
		}

		/// <summary>
		/// DB 작업을 요청할 쿼리문 </summary>
		void SetQuery(const WCHAR* pszQuery) { m_strQuery = pszQuery; }

		/// <summary>
		///  권한 지정 </summary>
		void SetConnectionMode(ConnectModeEnum nMode) { m_pConnection->PutMode(nMode); }

		/// <summary>
		/// 명시적 트랜잭션 사용. bAutoCommit이 false인 경우 명시적으로 커밋이나 롤백을 해야한다. </summary>
		void SetAutoCommit(const bool bAutoCommit)
		{
			m_bAutoCommit = bAutoCommit;

			if (m_bAutoCommit == false) {
				m_bCanCommitTransaction = false;
			}
		}

		/// <summary>
		/// 자동 커밋 가능 여부 </summary>
		bool CanAutoCommit() { return m_bAutoCommit; }

		/// <summary>
		/// 트랜잭션을 건다 </summary>
		void BeginTransaction()
		{
			try
			{
				m_pConnection->BeginTrans();
			}
			catch (_com_error &e)
			{
				dump_com_error(e);
				dump_user_error();
				return;
			}
		}

		/// <summary>
		/// 커밋. 트랜잭션을 걸었을 때 사용 </summary>
		void CommitTransaction()
		{
			try
			{
				m_pConnection->CommitTrans();
			}
			catch (_com_error &e)
			{
				dump_com_error(e);
				dump_user_error();
				return;
			}
		}

		/// <summary>
		/// 롤백. 트랜잭션을 걸었을 때 사용 </summary>
		void RollbackTransaction()
		{
			try
			{
				m_pConnection->RollbackTrans();
			}
			catch (_com_error &e)
			{
				dump_com_error(e);
				dump_user_error();
				return;
			}
		}

		/// <summary>
		/// 쿼리 작업 성공 여부 </summary>
		bool IsSuccess()
		{
			if (m_bCanGetParamGetFiled == false)
			{
				dump_user_error();

				m_strQuery.erase();
				m_strCommand.erase();
				m_strColumnName.erase();
				m_strParameterName.erase();
			}

			return m_bCanGetParamGetFiled;
		}

		bool CanGetParamGetFiled() { return m_bCanGetParamGetFiled; }

		void SetCommitTransaction() { m_bCanCommitTransaction = true; }

		bool CanCommitTransaction() { return m_bCanCommitTransaction; }

		/// <summary>
		/// 로그 </summary>
		void LOG(WCHAR* format, ...)
		{
			wchar_t szBuffer[1024] = { 0, };

			va_list ap;
			va_start(ap, format);
			vswprintf_s(szBuffer, 1024, format, ap);
			va_end(ap);

			wprintf(L"%s\n", szBuffer);
		}

		/// <summary>
		/// ADO 시스템 에러와 관련된 로그 출력 </summary>
		void dump_com_error(const _com_error& e)
		{
			m_bCanGetParamGetFiled = true;

			// 데이터를 가져올 수 없는 상황이므로 연결을 끊어버린다
			if (e.Error() == 0X80004005) {
				Close();
			}

			if (m_Config.CanWriteErrorLog())
			{
				LOG(L"Code = %08lX   Code meaning = %s", e.Error(), e.ErrorMessage());
				LOG(L"Source = %s", (LPCTSTR)e.Source());
				LOG(L"Desc = %s", (LPCTSTR)e.Description());
			}
		}

		/// <summary>
		/// ADO를 사용 에러와 관련된 로그 출력 </summary>
		void dump_user_error()
		{
			m_bCanGetParamGetFiled = true;

			if (m_Config.CanWriteErrorLog())
			{
				if (!m_strQuery.empty()) {
					LOG(L"SQLQuery[%s]", m_strQuery.c_str());
				}

				if (!m_strCommand.empty()) {
					LOG(L"Command[%s]", m_strCommand.c_str());
				}

				if (!m_strColumnName.empty()) {
					LOG(L"Column[%s]", m_strColumnName.c_str());
				}

				if (!m_strParameterName.empty()) {
					LOG(L"Paramter[%s]", m_strParameterName.c_str());
				}
			}
		}

		/// <summary>
		/// 필드 개수를 조회 
		/// <param name="nValue"> 필드 개수 </param>
		/// <returns> 성공(true) 실패(false) </returns> 
		/// </summary>
		bool GetFieldCount(OUT INT32& nValue)
		{
			m_strCommand = L"GetFieldCount()";

			try
			{
				nValue = m_pRecordset->GetFields()->GetCount();;
			}
			catch (_com_error &e)
			{
				dump_com_error(e);
				return false;
			}

			return true;
		}

		/// <summary>
		/// 다음 레코드로 이동 
		/// <returns> 성공(true) 실패(false) </returns>
		/// </summary>
		bool MoveNext()
		{
			m_strCommand = L"MoveNext()";

			try
			{
				m_pRecordset->MoveNext();
			}
			catch (_com_error &e)
			{
				dump_com_error(e);
				return false;
			}

			return true;
		}

		/// <summary>
		/// 쿼리에서 얻은 레코드의 끝인가?
		/// <returns> 끝이면 true,  다음에 레코드가 있다면 false </returns> 
		/// </summary>
		bool GetEndOfFile()
		{
			m_strCommand = _T("GetEndOfFile()");
			bool bEndOfFile = true;

			try
			{
				VARIANT_BOOL vbEnd = m_pRecordset->GetEndOfFile();

				if (vbEnd == 0) {
					bEndOfFile = false;
				}
			}
			catch (_com_error &e)
			{
				dump_com_error(e);
			}

			return bEndOfFile;
		}

		/// <summary>
		/// 다음 레코드셋으로 이동. 레코드셋은 쿼리문에서 select 문을 여러개 사용하는 경우 레코드셋이 복수개가 된다 
		/// <returns> 다음 레코드셋이 있다면 true,  끝이면 false </returns> 
		/// </summary>
		bool NextRecordSet()
		{
			m_strCommand = L"NextRecordSet()";

			_variant_t variantRec;
			variantRec.intVal = 0;

			try
			{
				m_pRecordset = m_pRecordset->NextRecordset((_variant_t*)&variantRec);
			}
			catch (_com_error &e)
			{
				dump_com_error(e);
				return false;
			}

			return true;
		}

		/// <summary>
		/// 프로시저 및 SQL Text을 실행한다. 부가기능 adCmdStoreProc, adCmdText처리 가능
		/// <param name="CommandType"> adCmdText는 slq문, adCmdStoredProc는 저장프로시저문 </param>
		/// <param name="OptionType"> adOptionUnspecified는 커맨드 실행 방법을 지정하지 않는다. </param>
		/// <returns> 성공(true) 실패(false) </returns> 
		/// </summary>
		bool Execute(CommandTypeEnum CommandType = adCmdStoredProc, ExecuteOptionEnum OptionType = adOptionUnspecified)
		{
			if (!m_bCanGetParamGetFiled) {
				return false;
			}

			try
			{
				if (m_pConnection->GetState() == adStateClosed && m_Config.IsCanRetryConnection())
				{
					m_strCommand = L"RetryOpen()";  	//재연결 시도

					if (RetryOpen() == false)
					{
						return false;
					}
					else
					{
						m_bCanGetParamGetFiled = true;
					}
				}

				m_strCommand = L"Execute()";

				m_pCommand->CommandType = CommandType;
				m_pCommand->CommandText = m_strQuery.c_str();

				if (m_Config.GetConnectionTimeout() != 0) {
					m_pCommand->CommandTimeout = m_Config.GetConnectionTimeout();
				}

				m_pRecordset = m_pCommand->Execute(NULL, NULL, OptionType);
			}
			catch (_com_error &e)
			{
				dump_com_error(e);
				return false;
			}

			return true;
		}

		/// <summary>
		/// 정수/실수/날짜시간 필드값을 읽는다.
		/// <param name="szName"> 컬럼 이름 </param>
		/// <param Value="Value"> 읽어 온 값 </param>
		/// <returns> 성공(true) 실패(false). 읽은 값이 null이면 실패를 리턴한다. </returns> 
		/// </summary>
		template<typename T> bool GetFieldValue(const WCHAR* szName, OUT T& Value)
		{
			m_strCommand = L"GetFieldValue(T)";
			m_strColumnName = szName;

			try {
				auto vFieldValue = m_pRecordset->GetCollect(szName);

				switch (vFieldValue.vt)
				{
				case VT_BOOL:		// bool
				case VT_I1:			// BYTE WORD
				case VT_I2:			// INT16
				case VT_UI1:		// UCHAR
				case VT_I4:			// DWORD
				case VT_DECIMAL:	// INT64
				case VT_R8:			// float double
				case VT_DATE:
					Value = vFieldValue;
					break;
				case VT_NULL:
				case VT_EMPTY:
					m_strColumnName += _T(" null value");
					dump_user_error();
					return FALSE;
				default:
					WCHAR sz[10] = { 0, };
					m_strColumnName += L" type error(vt = ";
					m_strColumnName += _itow_s(vFieldValue.vt, sz, 10);
					m_strColumnName += L" ) ";
					m_bCanGetParamGetFiled = false;
					return FALSE;
				}
			}
			catch (_com_error &e) {
				dump_com_error(e);
				return FALSE;
			}
			return TRUE;
		}

		/// <summary>
		/// 문자형 필드값을 읽는다.
		/// <param name="szName"> 컬럼 이름 </param>
		/// <param Value="pszValue"> 읽어 온 문자열 </param>
		/// <param Value="nSize"> pszValue의 길이 </param>
		/// <returns> 성공(true) 실패(false). 읽은 값이 null이거나 버퍼가 작다면 실패를 리턴한다. </returns> 
		/// </summary>
		bool GetFieldValue(const WCHAR* szName, OUT WCHAR* pszValue, const UINT32 nSize)
		{
			m_strCommand = L"GetFieldValue(string)";

			m_strColumnName = szName;

			try
			{
				_variant_t vFieldValue = m_pRecordset->GetCollect(szName);


				if (vFieldValue.vt == VT_NULL || vFieldValue.vt == VT_EMPTY)
				{
					m_strColumnName += L" null value";
					return false;
				}
				else if (vFieldValue.vt != VT_BSTR)
				{
					m_strColumnName += L" nonbstr type";
					return false;
				}

				if (nSize < wcslen((WCHAR*)(_bstr_t(vFieldValue.bstrVal))))
				{
					m_strColumnName += L" string size overflow";
					return false;
				}

				wcscpy_s(pszValue, nSize, (wchar_t*)static_cast<_bstr_t>(vFieldValue.bstrVal));

			}
			catch (_com_error &e)
			{
				dump_com_error(e);
				return false;
			}

			return true;
		}

		/// <summary>
		/// binary 필드값을 읽는다.
		/// <param name="szName"> 컬럼 이름 </param>
		/// <param Value="pbyBuffer"> 바이너리 데이터를 담을 버퍼 </param>
		/// <param Value="nBufferSize"> 버퍼의 크기 </param>
		/// <param Value="nReadSize"> 버퍼에 담은 데이터의 크기 </param>
		/// <returns> 성공(true) 실패(false). 읽은 값이 null이거나 버퍼가 작다면 실패를 리턴한다. </returns> 
		/// </summary>
		bool GetFieldValue(const WCHAR* pszName, OUT BYTE* pbyBuffer, const INT32 nBufferSize, OUT INT32& nReadSize)
		{
			m_strCommand = L"GetFieldValue(binary)";
			m_strColumnName = pszName;

			try
			{
				_variant_t vFieldValue = m_pRecordset->GetCollect(pszName);

				if (vFieldValue.vt == VT_NULL)
				{
					m_strColumnName += L" null value";
					return false;
				}
				else if (vFieldValue.vt != (VT_ARRAY | VT_UI1))
				{
					m_strColumnName += L" nonbinary type";
					return false;
				}

				FieldPtr pField = m_pRecordset->Fields->GetItem(pszName);

				if (nBufferSize < pField->ActualSize || nBufferSize > 8060)
				{
					m_strColumnName += L" binary size overflow";
					dump_user_error();
					return false;
				}

				nReadSize = static_cast<int>(pField->ActualSize);

				BYTE * pData = nullptr;
				SafeArrayAccessData(vFieldValue.parray, (void HUGEP* FAR*)&pData);
				CopyMemory(pbyBuffer, pData, static_cast<size_t>(pField->ActualSize));
				SafeArrayUnaccessData(vFieldValue.parray);
			}
			catch (_com_error &e)
			{
				dump_com_error(e);
				return false;
			}

			return true;
		}


		/// <summary>
		/// 정수/실수/날짜시간 타입의 파라메터 생성
		/// null값의 파라메터 생성은 CreateNullParameter을 사용  </summary
		template<typename T> void CreateParameter(IN wchar_t* pszName, IN enum DataTypeEnum Type, IN enum ParameterDirectionEnum Direction, IN T rValue)
		{
			if (!IsSuccess()) {
				return;
			}

			m_strCommand = L"CreateParameter(T)";
			m_strParameterName = pszName;

			try
			{
				_ParameterPtr pParametor = m_pCommand->CreateParameter(pszName, Type, Direction, 0);
				m_pCommand->Parameters->Append(pParametor);
				pParametor->Value = static_cast<_variant_t>(rValue);
			}
			catch (_com_error &e)
			{
				dump_com_error(e);
			}

			return;
		}

		/// <summary>
		/// 정수/실수/날짜시간 타입의 null값 파라메터 생성 /// </summary
		void CreateNullParameter(IN wchar_t* pszName, IN enum DataTypeEnum Type, IN enum ParameterDirectionEnum Direction)
		{
			if (!m_bCanGetParamGetFiled) {
				return;
			}

			m_strCommand = L"CreateParameter(null)";
			m_strParameterName = pszName;

			try
			{
				_ParameterPtr pParametor(m_pCommand->CreateParameter(pszName, Type, Direction, 0));
				m_pCommand->Parameters->Append(pParametor);

				_variant_t vNull;
				vNull.ChangeType(VT_NULL);
				pParametor->Value = vNull;
			}
			catch (_com_error &e)
			{
				dump_com_error(e);
			}

			return;
		}

		/// <summary>
		/// 문자열 타입 파라메터 생성, 길이 변수는 최소 0보다 커야 한다. null값 생성은 wchar_t*에 NULL값을 넘긴다.
		/// </summary
		void CreateParameter(IN wchar_t* pszName, IN enum DataTypeEnum Type, IN enum ParameterDirectionEnum Direction,
			IN wchar_t* pValue, IN int nSize)
		{
			if (!m_bCanGetParamGetFiled) {
				return;
			}

			m_strCommand = L"CreateParameter(TCHAR)";
			m_strParameterName = pszName;

			_ASSERTE(nSize > 0 && "not allow 0 size!!");

			try
			{
				_ParameterPtr pParametor(m_pCommand->CreateParameter(pszName, Type, Direction, nSize));
				m_pCommand->Parameters->Append(pParametor);

				if (pValue == nullptr)
				{
					_variant_t vValue;
					vValue.vt = VT_NULL;
					pParametor->Value = vValue;
				}
				else
				{
					_variant_t vValue(pValue);
					pParametor->Value = vValue;
				}
			}
			catch (_com_error &e)
			{
				dump_com_error(e);
			}

			return;
		}

		/// <summary>
		/// binary 타입 파라메터 생성, 길이 변수는 최소 0보다 커야 한다. null값 생성은 BYTE*에 NULL값을 넘긴다. </summary
		void CreateParameter(IN wchar_t* pszName, IN enum DataTypeEnum Type, IN enum ParameterDirectionEnum Direction,
			IN BYTE* pValue, IN int nSize)
		{
			if (!m_bCanGetParamGetFiled) {
				return;
			}

			_ASSERTE(nSize > 0 && "not allow 0 size!!");

			m_strCommand = L"CreateParameter(binary)";
			m_strParameterName = pszName;

			try
			{
				_ParameterPtr pParametor(m_pCommand->CreateParameter(pszName, Type, Direction, nSize));
				m_pCommand->Parameters->Append(pParametor);

				_variant_t vBinary;
				SAFEARRAY FAR *pArray = nullptr;
				SAFEARRAYBOUND rarrayBound[1];

				if (pValue == nullptr)		//명시적 null이거나 값이 null이라면
				{
					vBinary.vt = VT_NULL;
					pParametor->Value = vBinary;
				}
				else
				{
					vBinary.vt = VT_ARRAY | VT_UI1;
					rarrayBound[0].lLbound = 0;
					rarrayBound[0].cElements = nSize;
					pArray = SafeArrayCreate(VT_UI1, 1, rarrayBound);

					for (long n = 0; n < nSize; ++n)
					{
						SafeArrayPutElement(pArray, &n, &(pValue[n]));
					}
					vBinary.parray = pArray;
					pParametor->AppendChunk(vBinary);
				}
			}
			catch (_com_error &e)
			{
				dump_com_error(e);
			}

			return;
		}


		/// <summary>
		/// 정수/실수/날짜시간 타입의 파라메터값 변경
		/// null값의 파라메터 변경은 UpdateNullParameter을 사용 </summary
		template<typename T>
		void UpdateParameter(IN wchar_t* pszName, IN T rValue)
		{
			if (!IsSuccess()) {
				return;
			}

			m_strCommand = L"Updatesarameter(T)";
			m_strParameterName = pszName;

			try
			{
				m_pCommand->Parameters->GetItem(pszName)->Value = static_cast<_variant_t>(rValue);
			}
			catch (_com_error &e)
			{
				dump_com_error(e);
			}

			return;
		}

		/// <summary>
		/// 정수/실수/날짜시간 타입의 파라메터 값을 null로 변경 </summary
		void UpdateNullParameter(IN wchar_t* pszName)
		{
			if (!m_bCanGetParamGetFiled) {
				return;
			}

			m_strCommand = L"UpdateNullParameter(null)";
			m_strParameterName = pszName;

			try
			{
				m_pCommand->Parameters->GetItem(pszName)->Value.ChangeType(VT_NULL);
			}
			catch (_com_error &e)
			{
				dump_com_error(e);
			}

			return;
		}

		/// <summary>
		/// 문자열 타입 파라메터 변경, 길이 변수는 최소 0보다 커야 한다. null값 변경 TCHAR*에 NULL값을 넘긴다. </summary
		void UpdateParameter(IN wchar_t* pszName, IN wchar_t* pValue, IN int nSize)
		{
			if (!m_bCanGetParamGetFiled) {
				return;
			}

			_ASSERTE(nSize > 0 && "not allow 0 size!!");

			m_strCommand = L"UpdateParameter(WCHAR)";
			m_strParameterName = pszName;

			try
			{
				_variant_t vValue(pValue);

				if (pValue == nullptr)
				{
					vValue.vt = VT_NULL;
				}

				m_pCommand->Parameters->GetItem(pszName)->Size = nSize;
				m_pCommand->Parameters->GetItem(pszName)->Value = vValue;
			}
			catch (_com_error &e)
			{
				dump_com_error(e);
			}

			return;
		}

		/// <summary>
		/// binary 타입 파라메터 변경, 길이 변수는 최소 0보다 커야 한다. null값 변경 BYTE*에 NULL값을 넘긴다. </summary
		void UpdateParameter(IN wchar_t* pszName, IN BYTE* pValue, IN int nSize)
		{
			if (!m_bCanGetParamGetFiled) {
				return;
			}

			_ASSERTE(nSize > 0 && "not allow 0 size!!");

			m_strCommand = L"UpdateParameter(binary)";
			m_strParameterName = pszName;

			try
			{
				_ParameterPtr pParametor(m_pCommand->Parameters->GetItem(pszName));
				pParametor->Size = nSize;

				_variant_t vBinary;
				SAFEARRAY FAR *pArray = nullptr;
				SAFEARRAYBOUND rarrayBound[1];

				if (pValue == nullptr)
				{
					vBinary.vt = VT_NULL;
					rarrayBound[0].lLbound = 0;
					rarrayBound[0].cElements = 0;
					pParametor->Value = vBinary;
				}
				else
				{
					vBinary.vt = VT_ARRAY | VT_UI1;
					rarrayBound[0].lLbound = 0;
					rarrayBound[0].cElements = nSize;
					pArray = SafeArrayCreate(VT_UI1, 1, rarrayBound);

					for (long n = 0; n < nSize; ++n)
					{
						SafeArrayPutElement(pArray, &n, &(pValue[n]));
					}

					vBinary.parray = pArray;
					pParametor->AppendChunk(vBinary);
				}
			}
			catch (_com_error &e)
			{
				dump_com_error(e);
			}

			return;
		}

		/// <summary>
		/// 정수/실수/날짜시간 타입의 파라메터 값 읽기 </summary
		template<typename T>
		bool GetParameter(wchar_t* pszName, OUT T& Value)
		{
			if (!IsSuccess()) {
				return false;
			}

			m_bCanGetParamGetFiled = false;

			m_strCommand = L"GetParameter(T)";
			m_strParameterName = pszName;

			try
			{
				_variant_t& vFieldValue = m_pCommand->Parameters->GetItem(pszName)->Value;

				switch (vFieldValue.vt)
				{
				case VT_BOOL:	//bool
				case VT_I1:
				case VT_I2:		//BYTE WORD
				case VT_UI1:
				case VT_I4:		//DWORD
				case VT_DECIMAL: //INT64
				case VT_R8:	//float double
				case VT_DATE:
					Value = vFieldValue;
					break;
				case VT_NULL:
				case VT_EMPTY:
					m_strColumnName += L" null value";
					dump_user_error();
					return false;
				default:
					wchar_t sz[7] = { 0, };
					m_strParameterName += L" type error(vt = ";
					m_strParameterName += _itow_s(vFieldValue.vt, sz, 10);
					m_strParameterName += L" ) ";
					m_bCanGetParamGetFiled = false;
					return false;
				}
			}
			catch (_com_error &e)
			{
				dump_com_error(e);
				return false;
			}

			m_bCanGetParamGetFiled = true;

			return m_bCanGetParamGetFiled;
		}

		/**
		\remarks	문자형 파라메터값을 읽는다.
		\par		읽은 값이 null이거나 버퍼가 작다면 실패를 리턴한다.
		\param		읽은 문자을 담을 버퍼의 크기
		\return		성공(TRUE) 실패(FLASE)
		*/
		/// <summary>
		/// 
		/// </summary
		bool GetParameter(IN wchar_t* pszName, OUT wchar_t* pValue, IN unsigned int nSize)
		{
			if (!m_bCanGetParamGetFiled) {
				return false;
			}

			m_bCanGetParamGetFiled = false;

			m_strCommand = L"GetParameter(wchar_t*)";
			m_strParameterName = pszName;

			try
			{
				_variant_t vFieldValue = m_pCommand->Parameters->GetItem(pszName)->Value;

				if (vFieldValue.vt == VT_NULL || vFieldValue.vt == VT_EMPTY)
				{
					m_strParameterName += L" null value";
					return false;
				}
				else if (vFieldValue.vt != VT_BSTR)
				{
					m_strParameterName += L" nonString Type";
					return false;
				}
				else if (nSize < wcslen((wchar_t*)(_bstr_t(vFieldValue.bstrVal))))
				{
					m_strParameterName += L" string size overflow";
					return false;
				}

				wcscpy_s(pValue, nSize, (wchar_t*)(_bstr_t)vFieldValue);
			}
			catch (_com_error &e)
			{
				dump_com_error(e);
				return false;
			}

			m_bCanGetParamGetFiled = true;
			return m_bCanGetParamGetFiled;
		}

		/**
		\remarks	바이너리형 파라메터값을 읽는다.
		\par		읽은 값이 null이거나 버퍼가 작다면 실패를 리턴한다.
		\param		읽은 문자을 담을 버퍼의 크기, 읽은 버퍼의 크기
		\return		성공(TRUE) 실패(FLASE)
		*/
		/// <summary>
		/// 
		/// </summary
		bool GetParameter(IN wchar_t* pszName, OUT BYTE* pBuffer, IN int inSize, OUT int& outSize)
		{
			if (!m_bCanGetParamGetFiled) {
				return false;
			}

			m_bCanGetParamGetFiled = false;

			m_strCommand = L"GetParameter(binary)";
			m_strParameterName = pszName;

			try
			{
				_variant_t vFieldValue = m_pCommand->Parameters->GetItem(pszName)->Value;

				if (VT_NULL == vFieldValue.vt)
				{
					m_strParameterName += L" null value";
					return false;
				}
				else if ((VT_ARRAY | VT_UI1) != vFieldValue.vt)
				{
					m_strParameterName += L" nonbinary type";
					return false;
				}

				int ElementSize = vFieldValue.parray->rgsabound[0].cElements;

				if (ElementSize > inSize || inSize > 8060)
				{
					m_strParameterName += L" size overflow";
					return false;
				}

				BYTE * pData = nullptr;
				SafeArrayAccessData(vFieldValue.parray, (void HUGEP* FAR*)&pData);
				CopyMemory(pBuffer, pData, ElementSize);
				SafeArrayUnaccessData(vFieldValue.parray);
				outSize = vFieldValue.parray->rgsabound[0].cElements;
			}
			catch (_com_error &e)
			{
				dump_com_error(e);
				return false;
			}

			m_bCanGetParamGetFiled = true;
			return m_bCanGetParamGetFiled;
		}


	protected:
		_ConnectionPtr m_pConnection;
		_RecordsetPtr m_pRecordset;
		_CommandPtr m_pCommand;

		/// <summary> 자동 커밋 여부. false인 경우 커밋이나 롤백을 명시적으로 사용해야 한다 </summary>
		bool m_bAutoCommit;

		/// <summary> DB 설정 정보 </summary>
		DBConfig m_Config;

		std::wstring m_strQuery;

		/// <summary> Ado 파리미터나 테이블의 필드를 읽을 수 있는지 여부 </summary>
		bool m_bCanGetParamGetFiled;

		/// <summary> 커밋을 할 수 있는지 여부 </summary>
		bool m_bCanCommitTransaction;

		std::wstring m_strCommand;
		std::wstring m_strColumnName;
		std::wstring m_strParameterName;


		AdoDB(const AdoDB&);
		AdoDB& operator= (const AdoDB&);
	};



	

	class DBManager
	{
		enum{MAX_ARRAY_SIZE=20};
	public:
		explicit DBManager( DBConfig& dboconfig )
		{
			m_bSuccessConnection = true;
			int MaxConnectionPoolCount = dboconfig.GetMaxConnectionPool();

			_ASSERTE( MaxConnectionPoolCount <= MAX_ARRAY_SIZE );

			for( int i = 0; i < MaxConnectionPoolCount; ++i )
			{
				m_pAdoStack[i] = new AdoDB(dboconfig);

				if( m_pAdoStack[i]->Open() == false )
				{
					m_bSuccessConnection = false;
					break;
				}
			}

			m_nTopPos = m_nMaxAdo = MaxConnectionPoolCount - 1;
		}

		// 연결 성공 여부
		bool IsSuccessConnection() { return m_bSuccessConnection; }
	
		void PutDB( AdoDB* pAdo )
		{
			ScopedLock lock(m_Lock);
			
			_ASSERTE( m_nTopPos < m_nMaxAdo );

			m_pAdoStack[ ++m_nTopPos ] = pAdo;
			return;
		}

		AdoDB* GetDB()
		{
			ScopedLock lock(m_Lock);

			_ASSERTE( m_nTopPos >= 0 );
			
			return m_pAdoStack[ m_nTopPos-- ];
		}

	private:
		int m_nTopPos;
		int m_nMaxAdo;
		bool m_bSuccessConnection;	 // 연결 성공 여부
	
		AdoDB* m_pAdoStack[MAX_ARRAY_SIZE];
		CSSpinLockWin32 m_Lock;
	};


	/**
	\brief		객체 생성시 커넥션풀로부터 ADO객체를 얻은 후 소멸시 ADO객체를 커넥션풀로 돌려준다.
	\par		부가기능 명시적 트랜잭션
	\author		김영찬
	*/
	class CScopedAdo
	{
	public:
		explicit CScopedAdo(AdoDB* &pAdo, DBManager* pAdoManager, bool bAutoCommit = false)
			:m_pAdoManager(pAdoManager)
		{
			m_pAdo = pAdoManager->GetDB();
			pAdo = m_pAdo;
			pAdo->SetAutoCommit(bAutoCommit);

			if (bAutoCommit == false)
			{
				pAdo->BeginTransaction();
			}
		}

		~CScopedAdo()
		{
			if (m_pAdo->CanAutoCommit() == false)
			{
				if (m_pAdo->CanGetParamGetFiled() && m_pAdo->CanCommitTransaction())
				{
					m_pAdo->CommitTransaction();
				}
				else
				{
					m_pAdo->RollbackTransaction();
				}
			}

			m_pAdo->Init();
			m_pAdo->Release();
			m_pAdoManager->PutDB(m_pAdo);
		}

	private:
		DBManager* m_pAdoManager;
		AdoDB* m_pAdo;
	};
}
