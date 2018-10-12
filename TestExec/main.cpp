#include "stdafx.h"
#include <iostream>

#include "TestExec.h"
#include "TestInsertSP.h"

int _tmain(int argc, _TCHAR* argv[])
{
	setlocale(LC_ALL, "");

	TestExec();
	TestInsertSP();

	getchar();
	return 0;
}