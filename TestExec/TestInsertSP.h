#include <iostream>
#include "..\include\AdoManager.h"


/*
CREATE TABLE [dbo].[Test_Temp2](
	[ID]		[nvarchar](50)	NOT NULL,
	[UserCode]	[int]			NOT NULL,
	[Lv]		[int]			NOT NULL,
	[Money]		[bigint]		NOT NULL
) ON [PRIMARY]
*/

void TestInsertSP()
{
	setlocale(LC_ALL, "");
	
	AsyncAdodb::DBConfig config;
	config.Setting( L"gunz2db\\gunz2_db", 
						L"dev", 
						L"dev", 
						L"G2_GAMEDB", 
						3, 
						true, 
						3
					);

	auto pDBmanager = new AsyncAdodb::DBManager( config );
		
	// SQL Query - insert: auto commit
	{
		AsyncAdodb::AdoDB* pAdo = nullptr;
		AsyncAdodb::CScopedAdo scopedado( pAdo, pDBmanager, true );

		pAdo->SetQuery( L"Insert Into Test_Temp2 Values( 'jacking1', 1111 , 1, 100)" );
		pAdo->Execute(::adCmdText);
		
		if( !pAdo->IsSuccess() ) 
		{
			std::wcout << L"쿼리문 실패" << std::endl;
			return;
		}
		else
		{
			std::wcout << L"쿼리문 성공" << std::endl;
		}
	}

	// SQL Query - insert
	{
		AsyncAdodb::AdoDB* pAdo = nullptr;
		AsyncAdodb::CScopedAdo scopedado( pAdo, pDBmanager, false );

		pAdo->SetQuery( L"Insert Into Test_Temp2 Values( 'jacking2', 1112, 1, 100 )" );
		pAdo->Execute(adCmdText);
		
		if( !pAdo->IsSuccess() ) 
		{
			std::wcout << L"쿼리문 실패" << std::endl;
			return;
		}
		else
		{
			std::wcout << L"쿼리문 성공" << std::endl;
		}

		pAdo->SetCommitTransaction();
	}

	// SQL Query - select
	{
		AsyncAdodb::AdoDB* pAdo = nullptr;
		AsyncAdodb::CScopedAdo scopedado( pAdo, pDBmanager, true );

		pAdo->SetQuery(L"SELECT ID, Code FROM Test_Temp2 WHERE ID='jacking3'");
		pAdo->Execute(adCmdText);

		if( !pAdo->IsSuccess() )
		{
			std::wcout << L"select 쿼리문 실패" << std::endl;
			return;
		}

		WCHAR szID[16] = {0,};
		int nUserCode = 0;

		if( !pAdo->GetEndOfFile() )
		{
			pAdo->GetFieldValue(_T("ID"), szID, 16);
			pAdo->GetFieldValue(_T("UserCode"), nUserCode);
		}
		else
		{
			std::wcout << L"jacking3는 없습니다" << std::endl;
			return;
		}

		std::wcout << L"ID : " << szID << std::endl;
		std::wcout << L"UserCode : " << nUserCode << std::endl;
	}

	delete pDBmanager;
}