#include "TaskSchedulerFactory.h"
//  Include the task header file.

#pragma comment(lib, "taskschd.lib")
#pragma comment(lib, "comsupp.lib")
#pragma comment(lib, "credui.lib")

std::wstring TaskSchedulerFactory::m_Author = L"Wilson";

int TaskSchedulerFactory::Initialize()
{
    HRESULT hr = CoInitializeEx(NULL, COINIT_MULTITHREADED);
    if (FAILED(hr))
    {
        printf("\nCoInitializeEx failed: %x", hr);
        return INITIALIZE_ERROR;
    }

    //  Set general COM security levels.
    hr = CoInitializeSecurity(
        NULL,
        -1,
        NULL,
        NULL,
        RPC_C_AUTHN_LEVEL_PKT_PRIVACY,
        RPC_C_IMP_LEVEL_IMPERSONATE,
        NULL,
        0,
        NULL);

    if (FAILED(hr))
    {
        printf("\nCoInitializeSecurity failed: %x", hr);
        CoUninitialize();
        return INITIALIZE_SECURITY_ERROR;
    }

    return SUCCESS;
}

int TaskSchedulerFactory::CreateTaskService(LPVOID* pServiceTmp, ITaskFolder** pRootFolder)
{        
    HRESULT hr = CoCreateInstance(CLSID_TaskScheduler,
        NULL,
        CLSCTX_INPROC_SERVER,
        IID_ITaskService,
        pServiceTmp);
    if (FAILED(hr))
    {
        printf("Failed to create an instance of ITaskService: %x", hr);
        CoUninitialize();
        return CREATE_SERVICE_ERROR;
    }    
    ITaskService* pService = reinterpret_cast<ITaskService*>(*pServiceTmp);;
    //  Connect to the task service.
    hr = pService->Connect(_variant_t(), _variant_t(),
        _variant_t(), _variant_t());
    if (FAILED(hr))
    {
        printf("ITaskService::Connect failed: %x", hr);
        pService->Release();
        CoUninitialize();
        return CONNECT_SERVICE_ERROR;
    }

    //  ------------------------------------------------------
//  Get the pointer to the root task folder.  This folder will hold the
//  new task that is registered.    
    hr = pService->GetFolder(_bstr_t(L"\\"), pRootFolder);
    if (FAILED(hr))
    {
        printf("Cannot get Root folder pointer: %x", hr);
        pService->Release();
        CoUninitialize();
        return GET_ROOT_TASK_FOLDER_ERROR;
    }

    return SUCCESS;
}


int TaskSchedulerFactory::CreateTaskDefinition(ITaskService* pService, ITaskDefinition** pTask, IRegistrationInfo** pRegInfo)
{
    //  Create the task definition object to create the task.       
    HRESULT hr = pService->NewTask(0, pTask);    
    if (FAILED(hr))
    {
        printf("Failed to CoCreate an instance of the TaskService class: %x", hr);        
        CoUninitialize();
        return COCREATE_TASKSERVICE_INSTANCE_ERROR;
    }

    //  Get the registration info for setting the identification.    
    hr = (*pTask)->get_RegistrationInfo(pRegInfo);
    if (FAILED(hr))
    {
        printf("\nCannot get identification pointer: %x", hr);                
        CoUninitialize();
        return GET_REGISTRATION_ERROR;
    }
  
    BSTR bstr = SysAllocString(m_Author.c_str());
    hr = (*pRegInfo)->put_Author(bstr);
    (*pRegInfo)->Release();
    SysFreeString(bstr);
    if (FAILED(hr))
    {
        printf("\nCannot put identification info: %x", hr);
        CoUninitialize();
        return PUT_AUTHOR_ERROR;
    }

    return SUCCESS;
}

int TaskSchedulerFactory::SetupPrincipal(ITaskDefinition* pTask)
{
    //  ------------------------------------------------------
  //  Create the principal for the task - these credentials
  //  are overwritten with the credentials passed to RegisterTaskDefinition
    IPrincipal* pPrincipal = NULL;
    HRESULT hr = pTask->get_Principal(&pPrincipal);
    if (FAILED(hr))
    {
        printf("\nCannot get principal pointer: %x", hr);
        CoUninitialize();
        return GET_PRINCIPAL_ERROR;
    }

    //  Set up principal logon type to interactive logon
    hr = pPrincipal->put_LogonType(TASK_LOGON_INTERACTIVE_TOKEN);
    pPrincipal->Release();
    if (FAILED(hr))
    {
        printf("\nCannot put principal info: %x", hr);
        CoUninitialize();
        return PUT_PRINCIPAL_ERROR;
    }

    return SUCCESS;
}

int TaskSchedulerFactory::SetupTaskSettings(ITaskDefinition* pTask, bool wakeToRun, bool disallowStartIfOnBatteries, bool stopIfGoingOnBatteries)
{
    //  Create the settings for the task
    ITaskSettings* pSettings = NULL;
    HRESULT hr = pTask->get_Settings(&pSettings);
    if (FAILED(hr))
    {
        printf("\nCannot get settings pointer: %x", hr);                
        CoUninitialize();
        return SETUP_TASK_SETTING_ERROR;
    }

    //  Set setting values for the task.  
    hr = pSettings->put_StartWhenAvailable(VARIANT_TRUE);
    if (FAILED(hr))
    {
        printf("\nCannot put setting information: %x", hr);
        CoUninitialize();
        return SETUP_TASK_SETTING_ERROR;
    }

    hr = pSettings->put_WakeToRun(wakeToRun == true ? VARIANT_TRUE : VARIANT_FALSE);
    if (FAILED(hr))
    {
        printf("\nCannot put setting information: %x", hr);
        CoUninitialize();
        return SETUP_TASK_SETTING_ERROR;
    }

    hr = pSettings->put_DisallowStartIfOnBatteries(disallowStartIfOnBatteries == true ? VARIANT_TRUE : VARIANT_FALSE);
    if (FAILED(hr))
    {
        printf("\nCannot put setting information: %x", hr);
        CoUninitialize();
        return SETUP_TASK_SETTING_ERROR;
    }

    hr = pSettings->put_StopIfGoingOnBatteries(stopIfGoingOnBatteries == true ? VARIANT_TRUE : VARIANT_FALSE);
    pSettings->Release();
    if (FAILED(hr))
    {
        printf("\nCannot put setting information: %x", hr);
        CoUninitialize();
        return SETUP_TASK_SETTING_ERROR;
    }        

    //// Set the idle settings for the task.
    //IIdleSettings* pIdleSettings = NULL;
    //hr = pSettings->get_IdleSettings(&pIdleSettings);
    //pSettings->Release();
    //if (FAILED(hr))
    //{
    //    printf("\nCannot get idle setting information: %x", hr);
    //    CoUninitialize();
    //    return SETUP_TASK_SETTING_ERROR;
    //}    

    //BSTR pt5m = SysAllocString(L"PT5M");
    //hr = pIdleSettings->put_WaitTimeout(pt5m);
    //pIdleSettings->Release();
    //SysFreeString(pt5m);
    //if (FAILED(hr))
    //{
    //    printf("\nCannot put idle setting information: %x", hr);
    //    CoUninitialize();
    //    return 1;
    //}

    return SUCCESS;
}

int TaskSchedulerFactory::SetupTrigger(ITaskDefinition* pTask, uint32_t startAfterSeconds)
{
    //  Get the trigger collection to insert the time trigger.
    ITriggerCollection* pTriggerCollection = NULL;
    HRESULT hr = pTask->get_Triggers(&pTriggerCollection);
    if (FAILED(hr))
    {
        printf("\nCannot get trigger collection: %x", hr);                
        CoUninitialize();
        return SETUP_TRIGGER_ERROR;
    }

    //  Add the time trigger to the task.
    ITrigger* pTrigger = NULL;
    hr = pTriggerCollection->Create(TASK_TRIGGER_TIME, &pTrigger);
    pTriggerCollection->Release();
    if (FAILED(hr))
    {
        printf("\nCannot create trigger: %x", hr);
        CoUninitialize();
        return SETUP_TRIGGER_ERROR;
    }

    ITimeTrigger* pTimeTrigger = NULL;
    hr = pTrigger->QueryInterface(
        IID_ITimeTrigger, (void**)&pTimeTrigger);
    pTrigger->Release();
    if (FAILED(hr))
    {
        printf("\nQueryInterface call failed for ITimeTrigger: %x", hr);
        CoUninitialize();
        return SETUP_TRIGGER_ERROR;
    }

    hr = pTimeTrigger->put_Id(_bstr_t(L"Trigger1"));
    if (FAILED(hr))
        printf("\nCannot put trigger ID: %x", hr);

    //hr = pTimeTrigger->put_EndBoundary(_bstr_t(FormatSystemTime(startAfterSeconds + 60).c_str()));
    //if (FAILED(hr))
    //    printf("\nCannot put end boundary on trigger: %x", hr);

    //  Set the task to start at a certain time. The time 
    //  format should be YYYY-MM-DDTHH:MM:SS(+-)(timezone).
    hr = pTimeTrigger->put_StartBoundary(_bstr_t(FormatSystemTime(startAfterSeconds).c_str()));
    pTimeTrigger->Release();
    if (FAILED(hr))
    {
        printf("\nCannot add start boundary to trigger: %x", hr);
        CoUninitialize();
        return SETUP_TRIGGER_ERROR;
    }

    return SUCCESS;
}

int TaskSchedulerFactory::SetupAction(ITaskDefinition* pTask, ITaskFolder* pRootFolder, const std::wstring& taskName, const std::wstring& executablePath)
{
    //  ------------------------------------------------------
    //  Add an action to the task. This task will execute notepad.exe.     
    IActionCollection* pActionCollection = NULL;

    //  Get the task action collection pointer.
    HRESULT hr = pTask->get_Actions(&pActionCollection);
    if (FAILED(hr))
    {
        printf("\nCannot get Task collection pointer: %x", hr);
        CoUninitialize();
        return SETUP_ACTION_ERROR;
    }

    //  Create the action, specifying that it is an executable action.
    IAction* pAction = NULL;
    hr = pActionCollection->Create(TASK_ACTION_EXEC, &pAction);
    pActionCollection->Release();
    if (FAILED(hr))
    {
        printf("\nCannot create the action: %x", hr);
        CoUninitialize();
        return SETUP_ACTION_ERROR;
    }

    IExecAction* pExecAction = NULL;
    //  QI for the executable task pointer.
    hr = pAction->QueryInterface(
        IID_IExecAction, (void**)&pExecAction);
    pAction->Release();
    if (FAILED(hr))
    {
        printf("\nQueryInterface call failed for IExecAction: %x", hr);
        CoUninitialize();
        return SETUP_ACTION_ERROR;
    }

    //  Set the path of the executable to notepad.exe.
    hr = pExecAction->put_Path(_bstr_t(executablePath.c_str()));
    pExecAction->Release();
    if (FAILED(hr))
    {
        printf("\nCannot put action path: %x", hr);
        CoUninitialize();
        return SETUP_ACTION_ERROR;
    }

    LPCWSTR wszTaskName = L"ICVT task";
    //  ------------------------------------------------------
    //  Save the task in the root folder.
    IRegisteredTask* pRegisteredTask = NULL;
    hr = pRootFolder->RegisterTaskDefinition(
        _bstr_t(wszTaskName),
        pTask,
        TASK_CREATE_OR_UPDATE,
        _variant_t(),
        _variant_t(),
        TASK_LOGON_INTERACTIVE_TOKEN,
        _variant_t(L""),
        &pRegisteredTask);
    if (FAILED(hr))
    {
        printf("\nError saving the Task : %x", hr);
        CoUninitialize();
        return SETUP_ACTION_ERROR;
    }

    pRegisteredTask->Release();

    return SUCCESS;
}

int TaskSchedulerFactory::AddTask(const std::wstring& taskName, const std::wstring& executablePath, uint32_t startAfterSeconds, bool wakeupToRun, bool disallowStartIfOnBatteries, bool stopIfGoingOnBatteries)
{
    int err = Initialize();
    if (err != SUCCESS)
        return err;

    ITaskService* pService = NULL;
    ITaskFolder* pRootFolder = NULL;
    err = CreateTaskService((void**)&pService, &pRootFolder);

    //  If the same task exists, remove it.
    pRootFolder->DeleteTask(_bstr_t(taskName.c_str()), 0);

    //  Create the task definition object to create the task.   
    ITaskDefinition* pTask = NULL;
    IRegistrationInfo* pRegInfo = NULL;
    err = CreateTaskDefinition(pService, &pTask, &pRegInfo);
    pService->Release();

    if (err != SUCCESS)
    {
        pTask->Release();
        pRootFolder->Release();       
        return err;
    }   

    err = SetupPrincipal(pTask);
    if (err != SUCCESS)
    {
        pRootFolder->Release();
        pTask->Release();
        return err;
    }

    err = SetupTaskSettings(pTask, wakeupToRun, disallowStartIfOnBatteries, stopIfGoingOnBatteries);
    if (err != SUCCESS)
    {
        pRootFolder->Release();
        pTask->Release();
        return err;
    }

    err = SetupTrigger(pTask, startAfterSeconds);
    if (err != SUCCESS)
    {
        pRootFolder->Release();
        pTask->Release();
        return err;
    }
 
    err = SetupAction(pTask, pRootFolder, taskName, executablePath);
    if (err != SUCCESS)
    {
        pRootFolder->Release();
        pTask->Release();
        return err;
    }

    pRootFolder->Release();
    pTask->Release();
    CoUninitialize();

    return SUCCESS;
}

int TaskSchedulerFactory::DeleteTask(const std::wstring& taskName)
{
    int err = Initialize();
    if (err != SUCCESS)
        return err;

    ITaskService* pService = NULL;
    ITaskFolder* pRootFolder = NULL;
    err = CreateTaskService((void**)&pService, &pRootFolder);

    if (err != SUCCESS)
        return err;

    //  If the same task exists, remove it.
    HRESULT hr = pRootFolder->DeleteTask(_bstr_t(taskName.c_str()), 0);
    
    if (FAILED(hr) && hr != TASK_FILE_NOT_FOUND)
    {
        printf("\nError saving the Task : %x", hr);        
        err = DELETE_TASK_ERROR;
    }

    pService->Release();
    pRootFolder->Release();
    CoUninitialize();

    return err;
}

std::wstring TaskSchedulerFactory::FormatSystemTime(uint32_t afterSeconds)
{       
    time_t start = time(0) + afterSeconds;

    std::tm* ptm = std::localtime(&start);
    char buffer[32];    
    //2022-07-21T14:05:00
    std::strftime(buffer, 32, "%Y-%m-%dT%H:%M:%S", ptm);
    std::string str(buffer);
    std::wstring wstr(str.begin(), str.end());

    return wstr;
}