#include <windows.h>
#include <vector>
#include <string>
#include <algorithm>
#include <fstream>
#include <sstream>
#include <stdio.h>

// 如果 Windows SDK 版本过低，则手动定义缺失的常量
#ifndef DISPLAYCONFIG_PATH_PREFERRED_PRIMARY
#define DISPLAYCONFIG_PATH_PREFERRED_PRIMARY 0x00000004
#endif

// 日志记录函数 (可以选择保留或移除)
void LogToFile(const std::wstring& message) {
    std::wofstream logfile("DisplayCore.log", std::ios::app);
    if (logfile.is_open()) {
        logfile << message << std::endl;
    }
}

// 定义导出宏
#define DLLEXPORT extern "C" __declspec(dllexport)

typedef LONG(WINAPI* PFN_GET_DISPLAY_CONFIG_BUFFER_SIZES)(UINT32, UINT32*, UINT32*);
typedef LONG(WINAPI* PFN_QUERY_DISPLAY_CONFIG)(UINT32, UINT32*, DISPLAYCONFIG_PATH_INFO*, UINT32*, DISPLAYCONFIG_MODE_INFO*, DISPLAYCONFIG_TOPOLOGY_ID*);
typedef LONG(WINAPI* PFN_SET_DISPLAY_CONFIG)(UINT32, DISPLAYCONFIG_PATH_INFO*, UINT32, DISPLAYCONFIG_MODE_INFO*, UINT32);
typedef LONG(WINAPI* PFN_DISPLAY_CONFIG_GET_DEVICE_INFO)(DISPLAYCONFIG_DEVICE_INFO_HEADER*);

PFN_GET_DISPLAY_CONFIG_BUFFER_SIZES GetDisplayConfigBufferSizes_ = nullptr;
PFN_QUERY_DISPLAY_CONFIG QueryDisplayConfig_ = nullptr;
PFN_SET_DISPLAY_CONFIG SetDisplayConfig_ = nullptr;
PFN_DISPLAY_CONFIG_GET_DEVICE_INFO DisplayConfigGetDeviceInfo_ = nullptr;

void LoadCcdApi() {
    if (QueryDisplayConfig_ != nullptr) return;
    HMODULE user32 = GetModuleHandle(L"user32.dll");
    if (user32) {
        GetDisplayConfigBufferSizes_ = (PFN_GET_DISPLAY_CONFIG_BUFFER_SIZES)GetProcAddress(user32, "GetDisplayConfigBufferSizes");
        QueryDisplayConfig_ = (PFN_QUERY_DISPLAY_CONFIG)GetProcAddress(user32, "QueryDisplayConfig");
        SetDisplayConfig_ = (PFN_SET_DISPLAY_CONFIG)GetProcAddress(user32, "SetDisplayConfig");
        DisplayConfigGetDeviceInfo_ = (PFN_DISPLAY_CONFIG_GET_DEVICE_INFO)GetProcAddress(user32, "DisplayConfigGetDeviceInfo");
    }
}

LONG GetCurrentDisplayConfig(
    std::vector<DISPLAYCONFIG_PATH_INFO>& pathArray,
    std::vector<DISPLAYCONFIG_MODE_INFO>& modeArray) {
    LoadCcdApi();
    if (!QueryDisplayConfig_ || !GetDisplayConfigBufferSizes_) return ERROR_PROC_NOT_FOUND;

    UINT32 numPathArrayElements, numModeInfoArrayElements;
    LONG result = GetDisplayConfigBufferSizes_(QDC_ALL_PATHS, &numPathArrayElements, &numModeInfoArrayElements);
    if (result != ERROR_SUCCESS) return result;

    pathArray.resize(numPathArrayElements);
    modeArray.resize(numModeInfoArrayElements);

    return QueryDisplayConfig_(QDC_ALL_PATHS, &numPathArrayElements, pathArray.data(), &numModeInfoArrayElements, modeArray.data(), nullptr);
}

// 导出函数1：设置为单屏显示
DLLEXPORT LONG SetSingleDisplay(LPCWSTR targetDeviceName) {
    std::vector<DISPLAYCONFIG_PATH_INFO> pathArray;
    std::vector<DISPLAYCONFIG_MODE_INFO> modeArray;
    LONG result = GetCurrentDisplayConfig(pathArray, modeArray);
    if (result != ERROR_SUCCESS) return result;

    bool targetFound = false;
    // **【修复】** 引入一个布尔标志，确保只设置一次主显示器
    bool primaryAlreadySet = false;

    for (auto& path : pathArray) {
        DISPLAYCONFIG_SOURCE_DEVICE_NAME sourceName;
        sourceName.header.type = DISPLAYCONFIG_DEVICE_INFO_GET_SOURCE_NAME;
        sourceName.header.size = sizeof(sourceName);
        sourceName.header.adapterId = path.sourceInfo.adapterId;
        sourceName.header.id = path.sourceInfo.id;

        if (DisplayConfigGetDeviceInfo_(&sourceName.header) == ERROR_SUCCESS) {
             if (wcscmp(sourceName.viewGdiDeviceName, targetDeviceName) == 0) {
                targetFound = true;
                // **【修复】** 检查是否已经设置过主显示器
                if (!primaryAlreadySet) {
                    // 如果没有，则将此路径设为激活并设为首选主显示器
                    path.flags = DISPLAYCONFIG_PATH_ACTIVE | DISPLAYCONFIG_PATH_PREFERRED_PRIMARY;
                    primaryAlreadySet = true; // 标记已设置
                } else {
                    // 如果已经设置过，则仅激活此路径（处理冗余路径）
                    path.flags = DISPLAYCONFIG_PATH_ACTIVE;
                }
            } else {
                // 对于不匹配的设备，禁用其路径
                path.flags = 0;
            }
        } else {
            // 获取设备名失败的路径也禁用
            path.flags = 0;
        }
    }

    if (!targetFound) {
        return ERROR_NOT_FOUND;
    }
    
    // **【优化】** 在调用API前，清理掉所有无效的 source mode index 的路径
    // 这一步可以增加配置的纯净度，避免潜在问题
    pathArray.erase(std::remove_if(pathArray.begin(), pathArray.end(), 
        [](const DISPLAYCONFIG_PATH_INFO& path) {
            return path.sourceInfo.modeInfoIdx == DISPLAYCONFIG_PATH_MODE_IDX_INVALID;
        }), pathArray.end());


    return SetDisplayConfig_(pathArray.size(), pathArray.data(), modeArray.size(), modeArray.data(), SDC_APPLY | SDC_USE_SUPPLIED_DISPLAY_CONFIG | SDC_ALLOW_CHANGES);
}

// ... 省略其他未修改的函数 ...
// ... (SetExtendDisplays, SetCloneDisplays, SetExtendAllDisplays, DllMain) ...
// 您可以保持这些函数不变

// DLL入口点
BOOL APIENTRY DllMain(HMODULE hModule, DWORD ul_reason_for_call, LPVOID lpReserved) {
    if (ul_reason_for_call == DLL_PROCESS_ATTACH) {
        LoadCcdApi();
    }
    return TRUE;
}
