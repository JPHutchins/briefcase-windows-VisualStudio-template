#include <vcclr.h>

#include <Python.h>

#include "CrashDialog.h"

#include <fcntl.h>
#include <windows.h>

using namespace System;
using namespace System::Diagnostics;
using namespace System::IO;
using namespace System::Windows::Forms;

#define LOG_EXTENSION ".log"


static wchar_t* wstr(String^);
static String^ format_traceback(PyObject* type, PyObject* value, PyObject* traceback);

ref class LogWriter
{
private:
    String^ const filename;
    StreamWriter^ writer;
    FILE* file;
    int stderr_fileno;

public:
    LogWriter(String^ const filename) : filename{ filename } {
        writer = gcnew StreamWriter(filename);
    }

    void Write(String^ m) {
        writer->Write(m);
    }

    void WriteLine(String^ m) {
        writer->WriteLine(m);
    }

    void Flush(void) {
        writer->Flush();
    }

    /** @brief Capture stderr to the logfile. */
    void StartStdErrCapture(void) {
        // close the log file
        writer->Close();

        // save original stderr
        stderr_fileno = _dup(_fileno(stderr));

        // reopen the log file as the stderr destination
        FILE* _file = nullptr;
        errno_t const err = _wfreopen_s(&_file, wstr(filename), L"a", stderr);
        if (err != 0) {
            Console::Write("stderr redirect failed with err " + err + "\n");
            exit(-1);
        }
        else if (_file == nullptr) {
            Console::Write("Python system log file handle is nullptr\n");
            exit(-1);
        }
        int const fseek_err = fseek(_file, 0, SEEK_END);
        if (fseek_err != 0) {
            Console::Write("Failed to seek to end of file!\n");
            exit(-1);
        }
        file = _file;
    }

    /** @brief Stop capturing stderr to the logfile. */
    void StopStdErrCapture(void) {
        // restore original stderr
        int const dup2_err = _dup2(stderr_fileno, _fileno(stderr));
        if (dup2_err != 0) {
            Console::Write("Error resetting stderr!");
            exit(-1);
        }

        // reopen the log file for logging
        writer = gcnew StreamWriter(filename, true);
    }
};

ref class CrashDialogWriter {
public:
    LogWriter^ const log;

    CrashDialogWriter(LogWriter^ log) : log{ log } {}

    /** @brief Write an error message to the log file and stdout, then flush the log file. */
    void Write(String^ error_message) {
        log->Write(error_message);
        Console::Write(error_message);
        log->Flush();
    }
};

ref class ExitStatusExceptionHandler {
private:
    PyConfig * const config;
    CrashDialogWriter^ const crash_dialog;
    String^ const log_filename;

public:
    ExitStatusExceptionHandler(
        PyConfig* config,
        CrashDialogWriter^ const crash_dialog,
        String^ const log_filename
    ) 
        : config{ config }
        , crash_dialog{ crash_dialog } 
        , log_filename{ log_filename }
    {}

    void Handle(String^ error_message, PyStatus const * status) {
        crash_dialog->Write(error_message + "\n\t" + gcnew String(status->func) + ": " + gcnew String(status->err_msg) + "\n");
        crash_dialog->Write("\tSee " + log_filename + "\n");
        crash_dialog->log->StartStdErrCapture();
        PyConfig_Clear(config);
        Py_ExitStatusException(*status);
    }
};

int Main(array<String^>^ args) {
    int ret = 0;

    AttachConsole(ATTACH_PARENT_PROCESS);

    // Uninitialize the Windows threading model; allow apps to make
    // their own threading model decisions.
    CoUninitialize();

    // Get details of the app from app metadata
    FileVersionInfo^ const version_info = FileVersionInfo::GetVersionInfo(Application::ExecutablePath);

    String^ const log_folder = Environment::GetFolderPath(Environment::SpecialFolder::LocalApplicationData) + "\\" +
        version_info->CompanyName + "\\" +
        version_info->ProductName + "\\Logs";
    String^ const log_prefix = log_folder + "\\" + version_info->InternalName;
    String^ const python_system_error_log_name = log_folder + "\\python_system_error.log";
    String^ const current_log_name = log_prefix + LOG_EXTENSION;

    if (!Directory::Exists(log_folder)) {
        // If log folder doesn't exist, create it
        Directory::CreateDirectory(log_folder);
    } else {
        // If it does, rotate the logs in that folder.
        
        // Delete the log at index 9 if it exists
        String^ const evicted_log = log_prefix + "-9" + LOG_EXTENSION;
        if (File::Exists(evicted_log)) {
            File::Delete(evicted_log);
        }

        // - Move <app name>-8.log -> <app name>-9.log
        // - Move <app name>-7.log -> <app name>-8.log
        // - ...
        // - Move <app name>-2.log -> <app name>-3.log
        for (int log_index = 8; log_index >= 2; log_index--) {
            String^ const old_log_name = log_prefix + "-" + log_index + LOG_EXTENSION;
            if (File::Exists(old_log_name)) {
                File::Move(old_log_name, log_prefix + "-" + (log_index + 1) + LOG_EXTENSION);
            }
        }

        // Move <app name>.log -> <app name>-2.log
        if (File::Exists(current_log_name)) {
            File::Move(current_log_name, log_prefix + "-2" + LOG_EXTENSION);
        }
    }

    LogWriter^ const log = gcnew LogWriter(current_log_name);
    CrashDialogWriter^ const crash_dialog = gcnew CrashDialogWriter(log);

    log->WriteLine("Log started: " + DateTime::Now.ToString("yyyy-MM-dd HH:mm:ssZ"));

    // Preconfigure the Python interpreter;
    // This ensures the interpreter is in Isolated mode,
    // and is using UTF-8 encoding.
    log->WriteLine("PreInitializing Python runtime...");
    PyPreConfig pre_config;
    PyPreConfig_InitPythonConfig(&pre_config);
    pre_config.utf8_mode = 1;
    pre_config.isolated = 1;

    log->StartStdErrCapture();
    PyStatus status = Py_PreInitialize(&pre_config);
    log->StopStdErrCapture();

    if (PyStatus_Exception(status)) {
        crash_dialog->Write("Unable to pre-initialize Python runtime\n");
        crash_dialog->Write("See " + current_log_name + "\n");
        Py_ExitStatusException(status);
    }

    // Pre-initialize Python configuration
    PyConfig config;
    PyConfig_InitIsolatedConfig(&config);

    // Configure the Python interpreter:
    // Don't buffer stdio. We want output to appears in the log immediately
    config.buffered_stdio = 0;
    // Don't write bytecode; we can't modify the app bundle
    // after it has been signed.
    config.write_bytecode = 0;
    // Isolated apps need to set the full PYTHONPATH manually.
    config.module_search_paths_set = 1;

    ExitStatusExceptionHandler^ const py_exit_status = gcnew ExitStatusExceptionHandler(
        &config, crash_dialog, current_log_name
    );

    // Set the home for the Python interpreter
    String^ const python_home = Application::StartupPath;
    log->WriteLine("PythonHome: " + python_home);
    log->StartStdErrCapture();
    status = PyConfig_SetString(&config, &config.home, wstr(python_home));
    log->StopStdErrCapture();
    if (PyStatus_Exception(status)) {
        py_exit_status->Handle("Unable to set PYTHONHOME:", &status);
    }

    // Determine the app module name. Look for the BRIEFCASE_MAIN_MODULE
    // environment variable first; if that exists, we're probably in test
    // mode. If it doesn't exist, fall back to the MainModule key in the
    // main bundle.
    wchar_t* app_module_str;
    size_t size;
    _wdupenv_s(&app_module_str, &size, L"BRIEFCASE_MAIN_MODULE");
    String^ app_module_name;
    if (app_module_str) {
        app_module_name = gcnew String(app_module_str);
    } else {
        app_module_name = version_info->InternalName;
        app_module_str = wstr(app_module_name);
    }
    log->StartStdErrCapture();
    status = PyConfig_SetString(&config, &config.run_module, app_module_str);
    log->StopStdErrCapture();
    if (PyStatus_Exception(status)) {
        py_exit_status->Handle("Unable to set app module name:", &status);
    }

    // Read the site config
    log->StartStdErrCapture();
    status = PyConfig_Read(&config);
    log->StopStdErrCapture();
    if (PyStatus_Exception(status)) {
        py_exit_status->Handle("Unable to read site config:", &status);
    }

    // Set the full module path. This includes the stdlib, site-packages, and app code.
    log->WriteLine("PYTHONPATH:");
    // The .zip form of the stdlib
    String^ path = python_home + "\\python312.zip";
    log->WriteLine("- " + path);
    log->StartStdErrCapture();
    status = PyWideStringList_Append(&config.module_search_paths, wstr(path));
    log->StopStdErrCapture();
    if (PyStatus_Exception(status)) {
        py_exit_status->Handle("Unable to set .zip form of stdlib path:", &status);
    }

    // The unpacked form of the stdlib
    log->WriteLine("- " + python_home);
    log->StartStdErrCapture();
    status = PyWideStringList_Append(&config.module_search_paths, wstr(python_home));
    log->StopStdErrCapture();
    if (PyStatus_Exception(status)) {
        py_exit_status->Handle("Unable to set unpacked form of stdlib path:", &status);
    }

    // Add the app_packages path
    path = System::Windows::Forms::Application::StartupPath + "\\app_packages";
    log->WriteLine("- " + path);
    log->StartStdErrCapture();
    status = PyWideStringList_Append(&config.module_search_paths, wstr(path));
    log->StopStdErrCapture();
    if (PyStatus_Exception(status)) {
        py_exit_status->Handle("Unable to set app packages path:", &status);
    }

    // Add the app path
    path = System::Windows::Forms::Application::StartupPath + "\\app";
    log->WriteLine("- " + path);
    log->StartStdErrCapture();
    status = PyWideStringList_Append(&config.module_search_paths, wstr(path));
    log->StopStdErrCapture();
    if (PyStatus_Exception(status)) {
        py_exit_status->Handle("Unable to set app path:", &status);
    }

    log->WriteLine("Configure argc/argv...");
    wchar_t** argv = new wchar_t* [args->Length + 1];
    argv[0] = wstr(Application::ExecutablePath);
    for (int i = 0; i < args->Length; i++) {
        argv[i + 1] = wstr(args[i]);
    }
    log->StartStdErrCapture();
    status = PyConfig_SetArgv(&config, args->Length + 1, argv);
    log->StopStdErrCapture();
    if (PyStatus_Exception(status)) {
        py_exit_status->Handle("Unable to configure argc/argv:", &status);
    }

    log->WriteLine("Initializing Python runtime...");
    log->StartStdErrCapture();
    status = Py_InitializeFromConfig(&config);
    log->StopStdErrCapture();
    if (PyStatus_Exception(status)) {
        py_exit_status->Handle("Unable to initialize Python interpreter:", &status);
    }

    try {
        // Start the app module.
        //
        // From here to Py_ObjectCall(runmodule...) is effectively
        // a copy of Py_RunMain() (and, more  specifically, the
        // pymain_run_module() method); we need to re-implement it
        // because we need to be able to inspect the error state of
        // the interpreter, not just the return code of the module.
        log->WriteLine("Running app module: " + app_module_name);

        PyObject* const module = PyImport_ImportModule("runpy");
        if (module == NULL) {
            crash_dialog->Write("Could not import runpy module");
            exit(-2);
        }

        PyObject* const module_attr = PyObject_GetAttrString(module, "_run_module_as_main");
        if (module_attr == NULL) {
            crash_dialog->Write("Could not access runpy._run_module_as_main");
            exit(-3);
        }

        PyObject* const app_module = PyUnicode_FromWideChar(app_module_str, wcslen(app_module_str));
        if (app_module == NULL) {
            crash_dialog->Write("Could not convert module name to unicode");
            exit(-3);
        }

        PyObject* const method_args = Py_BuildValue("(Oi)", app_module, 0);
        if (method_args == NULL) {
            crash_dialog->Write("Could not create arguments for runpy._run_module_as_main");
            exit(-4);
        }

        // Print a separator to differentiate Python startup logs from app logs,
        // then flush the log and stdout/stderr to ensure all startup logs have been output.
        log->Write("---------------------------------------------------------------------------\n");
        log->Flush();
        fflush(stdout);
        fflush(stderr);

        // Invoke the app module
        PyObject const * const result = PyObject_Call(module_attr, method_args, NULL);

        // Print a separator to differentiate app logs from exit logs,
        // then flush the log and stdout/stderr to ensure all logs have been output.
        log->Write("---------------------------------------------------------------------------\n");
        log->Flush();
        fflush(stdout);
        fflush(stderr);

        if (result == NULL) {
            // Retrieve the current error state of the interpreter.
            PyObject* exc_type;
            PyObject* exc_value;
            PyObject* exc_traceback;
            PyErr_Fetch(&exc_type, &exc_value, &exc_traceback);
            PyErr_NormalizeException(&exc_type, &exc_value, &exc_traceback);

            if (exc_traceback == NULL) {
                crash_dialog->Write("Could not retrieve traceback");
                exit(-5);
            }

            if (PyErr_GivenExceptionMatches(exc_value, PyExc_SystemExit)) {
                PyObject* systemExit_code = PyObject_GetAttrString(exc_value, "code");
                if (systemExit_code == NULL) {
                    crash_dialog->Write("Could not determine exit code, setting to -10\n");
                    ret = -10;
                } else {
                    ret = (int)PyLong_AsLong(systemExit_code);
                }
            } else {
                ret = -6;
            }

            log->Write("Application will quit with exit code " + ret + "\n");

            if (ret != 0) {
                // Display stack trace in the crash dialog.
                crash_dialog->Write(format_traceback(exc_type, exc_value, exc_traceback));

                // Restore the error state of the interpreter.
                PyErr_Restore(exc_type, exc_value, exc_traceback);

                // Exit here so that Py_Finalize() does not also print the traceback
                exit(ret);
            }
        }
    }
    catch (Exception^ exception) {
        crash_dialog->Write("Python runtime error: " + exception->ToString());
        ret = -7;
    }

    Py_Finalize();
    return ret;
}

static inline wchar_t* wstr(String^ str)
{
    pin_ptr<const wchar_t> pinned = PtrToStringChars(str);
    return (wchar_t *) pinned;
}

/**
 * Convert a Python traceback object into a user-suitable string, stripping off
 * stack context that comes from this stub binary.
 *
 * If any error occurs processing the traceback, the error message returned
 * will describe the mode of failure.
 */
static String^ format_traceback(PyObject* type, PyObject* value, PyObject* traceback) {
    // Drop the top two stack frames; these are internal
    // wrapper logic, and not in the control of the user.
    for (int i = 0; i < 2; i++) {
        PyObject* inner_traceback = PyObject_GetAttrString(traceback, "tb_next");
        if (inner_traceback != NULL) {
            traceback = inner_traceback;
        }
    }

    // Format the traceback.
    PyObject* traceback_module = PyImport_ImportModule("traceback");
    if (traceback_module == NULL) {
        return "Could not import traceback";
    }

    PyObject* format_exception = PyObject_GetAttrString(traceback_module, "format_exception");
    PyObject* traceback_list = nullptr;
    if (format_exception && PyCallable_Check(format_exception)) {
        traceback_list = PyObject_CallFunctionObjArgs(format_exception, type, value, traceback, NULL);
    } else {
        return "Could not find 'format_exception' in 'traceback' module.";
    }
    if (traceback_list == NULL) {
        return "Could not format traceback.";
    }

    // Concatenate all the lines of the traceback into a single string
    PyObject* traceback_unicode = PyUnicode_Join(PyUnicode_FromString(""), traceback_list);

    // Convert the Python Unicode string into a UTF-16 Windows String.
    // It's easiest to do this by using the Python API to encode to UTF-8,
    // and then convert the UTF-8 byte string into UTF-16 using Windows APIs.
    Py_ssize_t size;
    const char* bytes = PyUnicode_AsUTF8AndSize(PyObject_Str(traceback_unicode), &size);

    System::Text::UTF8Encoding^ utf8 = gcnew System::Text::UTF8Encoding;
    String^ traceback_str = gcnew String(utf8->GetString((Byte*)bytes, (int)size));

    // Clean up the traceback string, removing references to the installed app location
    return traceback_str->Replace(Application::StartupPath, "");
}
