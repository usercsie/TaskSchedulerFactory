#pragma once

#include <windows.h>
#include <iostream>
#include <stdio.h>
#include <comdef.h>
#include <wincred.h>
#include <taskschd.h>
#include <time.h>


enum TaskSchedulerError
{
	SUCCESS = 0,
	INITIALIZE_ERROR,
	INITIALIZE_SECURITY_ERROR,
	CREATE_SERVICE_ERROR,
	CONNECT_SERVICE_ERROR,
	GET_ROOT_TASK_FOLDER_ERROR,
	COCREATE_TASKSERVICE_INSTANCE_ERROR,
	GET_REGISTRATION_ERROR,
	PUT_AUTHOR_ERROR,
	GET_PRINCIPAL_ERROR,
	PUT_PRINCIPAL_ERROR,
	SETUP_TASK_SETTING_ERROR,
	SETUP_TRIGGER_ERROR,
	SETUP_ACTION_ERROR,
	DELETE_TASK_ERROR,
};

#define TASK_FILE_NOT_FOUND 0x80070002

class TaskSchedulerFactory
{	
	static std::wstring m_Author;
private:
	static int Initialize();
	static int CreateTaskService(LPVOID* pService, ITaskFolder** pRootFolder);
	static int CreateTaskDefinition(_In_ ITaskService* pService, _Out_ ITaskDefinition** pTask, _Out_ IRegistrationInfo** pRegInfo);
	static int SetupPrincipal(ITaskDefinition* pTask);
	static int SetupTaskSettings(ITaskDefinition* pTask, bool wakeToRun, bool disallowStartIfOnBatteries, bool stopIfGoingOnBatteries);
	static int SetupTrigger(ITaskDefinition* pTask, uint32_t startAfterSeconds);
	static int SetupAction(ITaskDefinition* pTask, ITaskFolder* pRootFolder, const std::wstring &taskName, const std::wstring &executablePath);
	
public:
	static int AddTask(const std::wstring& taskName, const std::wstring& executablePath, uint32_t startAfterSeconds, bool wakeupToRun = true, bool disallowStartIfOnBatteries = false, bool stopIfGoingOnBatteries = false);
	static int DeleteTask(const std::wstring& taskName);
	static std::wstring FormatSystemTime(uint32_t afterSeconds);
};

